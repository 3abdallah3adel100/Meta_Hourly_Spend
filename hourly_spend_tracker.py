from __future__ import annotations

import json
import math
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials


# ============================================================
# BASIC CONFIG
# ============================================================

BASE_URL = "https://graph.facebook.com"

SHEET_TAB = (
    os.getenv("SHEET_TAB", "Spend Tracker").strip()
    or "Spend Tracker"
)

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ============================================================
# GOOGLE SHEET LAYOUT
# ============================================================

# Top KPI / view controls
VIEW_YEAR_CELL = "A5"
VIEW_MONTH_CELL = "D5"

LIVE_OVERALL_CELL = "G5"
SPEND_TODAY_CELL = "K5"
TODAY_BUDGET_CELL = "O5"
REMAINING_CELL = "S5"
LAST_UPDATE_CELL = "W5"


# Daily manual budget input
# A = Budget Date
# B = Overall Budget
# C:K = 9 Agents
# L = Notes
# M = Save Status
INPUT_HEADER_RANGE = "A8:M8"
INPUT_ROW_RANGE = "A9:M9"
INPUT_STATUS_CELL = "M9"


# Overall percentage matrix
# A13 = Day
# B13:Y13 = hourly headers
# B14:Y44 = 31 days x 24 hours
OVERALL_DAY_HEADER_CELL = "A13"
OVERALL_HOUR_HEADER_RANGE = "B13:Y13"
OVERALL_MATRIX_RANGE = "B14:Y44"


# Overall SPEND matrix — directly below the percentage table
# A47 = title
# A48 = Day
# B48:Y48 = hourly headers
# B49:Y79 = 31 days x 24 hours
SPEND_TITLE_CELL = "A47"
SPEND_DAY_HEADER_CELL = "A48"
SPEND_HOUR_HEADER_RANGE = "B48:Y48"
SPEND_MATRIX_RANGE = "B49:Y79"


# Agent percentage matrix — moved lower to make room for spend table
# A82 = section title
# A83 = selector label
# B83 = selected agent
# A85 = Day
# B85:Y85 = hourly headers
# B86:Y116 = 31 days x 24 hours
AGENT_TITLE_CELL = "A82"
AGENT_SELECTOR_LABEL_CELL = "A83"
AGENT_SELECTOR_CELL = "B83"
AGENT_DAY_HEADER_CELL = "A85"
AGENT_HOUR_HEADER_RANGE = "B85:Y85"
AGENT_MATRIX_RANGE = "B86:Y116"


# Live agents snapshot — moved lower
SNAPSHOT_TITLE_CELL = "A119"
SNAPSHOT_HEADER_RANGE = "A120:I120"
SNAPSHOT_RANGE = "A121:I129"


# Layout migration marker (Z is intentionally outside the visible matrix)
LAYOUT_VERSION_CELL = "Z1"
LAYOUT_VERSION = "V6_PRO_DESIGN_REPAIR"


# Internal Budget History
# AA:AM
HISTORY_HEADER_RANGE = "AA8:AM8"
HISTORY_DATA_START = 9


# Internal Raw Hourly Log
# AO:AX
RAW_HEADER_RANGE = "AO8:AX8"
RAW_DATA_START = 9


# ============================================================
# GENERIC HELPERS
# ============================================================

def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()

    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {name}"
        )

    return value


def split_values(raw: str) -> list[str]:
    return [
        value.strip()
        for value in re.split(
            r"[,;\n]+",
            str(raw or ""),
        )
        if value.strip()
    ]


def to_float(
    value: Any,
    default: float = 0.0,
) -> float:
    """
    Convert normal numbers AND Google Sheets percentage strings.

    Examples:
        59771.72  -> 59771.72
        "59,771.72" -> 59771.72
        0.6641 -> 0.6641
        "66.41%" -> 0.6641

    This fixes the matrix issue where gspread may return a formatted
    percentage like "66.41%" instead of the underlying numeric 0.6641.
    """
    try:
        raw = (
            str(value or "")
            .replace(",", "")
            .strip()
        )

        if not raw:
            return default

        is_percentage = raw.endswith("%")

        if is_percentage:
            raw = raw[:-1].strip()

        number = float(raw)

        if is_percentage:
            number /= 100.0

        return (
            number
            if math.isfinite(number)
            else default
        )

    except (TypeError, ValueError):
        return default


def parse_sheet_date(value: Any) -> str:
    raw = str(value or "").strip()

    if not raw:
        return ""

    candidates = [raw]

    if "T" in raw:
        candidates.append(
            raw.split("T", 1)[0]
        )

    if len(raw) >= 10:
        candidates.append(
            raw[:10]
        )

    for candidate in candidates:
        for fmt in (
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%m/%d/%Y",
            "%Y/%m/%d",
        ):
            try:
                return (
                    datetime.strptime(
                        candidate,
                        fmt,
                    )
                    .date()
                    .isoformat()
                )
            except ValueError:
                pass

    return raw


def parse_agent_header(
    value: Any,
) -> tuple[str, str] | None:
    """
    Expected:
        AA - Abdallah
        BM - Bassem
        HM - Etch

    The text before " - " is the code used to match
    Meta Ad Account names.
    """
    raw = str(value or "").strip()

    if not raw:
        return None

    if " - " in raw:
        code, name = raw.split(
            " - ",
            1,
        )

    elif "-" in raw:
        code, name = raw.split(
            "-",
            1,
        )

    else:
        code = raw
        name = raw

    code = code.strip().upper()
    name = name.strip() or code

    if not code:
        return None

    return code, name


def format_hour_12h(
    hour: int,
    minute: int = 0,
) -> str:
    """
    24h numeric hour -> readable 12-hour label.

    Examples:
        0,55  -> 12:55 AM
        1,55  -> 1:55 AM
        12,55 -> 12:55 PM
        17,55 -> 5:55 PM
    """
    dt = datetime(
        2000,
        1,
        1,
        hour,
        minute,
    )
    return dt.strftime(
        "%I:%M %p"
    ).lstrip("0")


def header_minute(
    value: Any,
    default: int = 0,
) -> int:
    """
    Preserve the minute already shown in an hourly header.
    Works with:
        17:55
        5:55 PM
        0:55
    """
    raw = str(value or "").strip()

    match = re.search(
        r":(\d{2})",
        raw,
    )

    if not match:
        return default

    try:
        minute = int(
            match.group(1)
        )
    except ValueError:
        return default

    return (
        minute
        if 0 <= minute <= 59
        else default
    )


def normalize_account_id(
    value: Any,
) -> str:
    return (
        str(value or "")
        .replace("act_", "")
        .strip()
    )


# ============================================================
# META HTTP HELPERS
# ============================================================

def request_json(
    url: str,
    params: dict[str, Any],
    attempts: int = 4,
) -> dict[str, Any]:

    last_error: Exception | None = None

    for attempt in range(
        1,
        attempts + 1,
    ):
        try:
            response = requests.get(
                url,
                params=params,
                timeout=60,
            )

            try:
                data = response.json()
            except ValueError:
                data = None

            if (
                response.ok
                and isinstance(data, dict)
            ):
                return data

            if isinstance(data, dict):
                error = data.get(
                    "error",
                    {},
                )

                detail = (
                    f"{error.get('message') or f'HTTP {response.status_code}'}"
                    f" | code={error.get('code')}"
                    f" | subcode={error.get('error_subcode')}"
                )

            else:
                detail = (
                    f"HTTP {response.status_code}: "
                    f"{response.text[:500]}"
                )

            if response.status_code not in {
                429,
                500,
                502,
                503,
                504,
            }:
                raise RuntimeError(
                    detail
                )

            last_error = RuntimeError(
                detail
            )

        except requests.RequestException as exc:
            last_error = exc

        except Exception as exc:
            last_error = exc

        if attempt < attempts:
            time.sleep(
                min(
                    8,
                    2 ** (attempt - 1),
                )
            )

    raise RuntimeError(
        str(
            last_error
            or "Unknown Meta API error"
        )
    )


def fetch_all_pages(
    url: str,
    params: dict[str, Any],
) -> list[dict[str, Any]]:

    rows: list[dict[str, Any]] = []

    next_url: str | None = url
    next_params: dict[str, Any] | None = params

    while next_url:

        payload = request_json(
            next_url,
            next_params or {},
        )

        rows.extend(
            payload.get(
                "data",
                [],
            )
        )

        next_url = (
            payload
            .get("paging", {})
            .get("next")
        )

        next_params = None

    return rows


# ============================================================
# META CLIENT
# ============================================================

class MetaClient:

    def __init__(self) -> None:

        self.api_version = (
            os.getenv(
                "META_API_VERSION",
                "v26.0",
            ).strip()
            or "v26.0"
        )

        self.business_ids = split_values(
            required_env(
                "BUSINESS_IDS"
            )
        )

        token_1 = required_env(
            "META_ACCESS_TOKEN"
        )

        token_2 = os.getenv(
            "META_ACCESS_TOKEN_2",
            "",
        ).strip()

        self.tokens: dict[str, str] = {
            "token_1": token_1,
        }

        if (
            token_2
            and token_2 != token_1
        ):
            self.tokens[
                "token_2"
            ] = token_2

        try:
            self.max_workers = max(
                1,
                min(
                    20,
                    int(
                        os.getenv(
                            "MAX_WORKERS",
                            "8",
                        )
                    ),
                ),
            )

        except ValueError:
            self.max_workers = 8


    def discover_accounts(
        self,
    ) -> list[dict[str, Any]]:
        """
        Discover accounts from every configured Business ID
        using both owned_ad_accounts and client_ad_accounts.

        If account appears more than once, it is deduped by account ID.
        """

        collected: dict[
            str,
            dict[str, Any],
        ] = {}

        errors: list[str] = []

        for business_id in self.business_ids:

            for edge in (
                "owned_ad_accounts",
                "client_ad_accounts",
            ):

                url = (
                    f"{BASE_URL}/"
                    f"{self.api_version}/"
                    f"{business_id}/"
                    f"{edge}"
                )

                any_success = False

                for (
                    token_key,
                    token,
                ) in self.tokens.items():

                    params = {
                        "fields": (
                            "id,"
                            "account_id,"
                            "name,"
                            "account_status,"
                            "currency"
                        ),
                        "access_token": token,
                        "limit": 500,
                    }

                    try:

                        rows = fetch_all_pages(
                            url,
                            params,
                        )

                        any_success = True

                        for row in rows:

                            account_id = (
                                normalize_account_id(
                                    row.get("id")
                                    or row.get(
                                        "account_id"
                                    )
                                )
                            )

                            if not account_id:
                                continue

                            if (
                                account_id
                                not in collected
                            ):

                                collected[
                                    account_id
                                ] = {
                                    "id": (
                                        f"act_"
                                        f"{account_id}"
                                    ),
                                    "account_name": str(
                                        row.get(
                                            "name"
                                        )
                                        or account_id
                                    ),
                                    "currency": str(
                                        row.get(
                                            "currency"
                                        )
                                        or "EGP"
                                    ),
                                    "token_key": (
                                        token_key
                                    ),
                                }

                    except Exception as exc:

                        clean = str(exc)

                        for secret in (
                            self.tokens.values()
                        ):
                            clean = clean.replace(
                                secret,
                                "[REDACTED]",
                            )

                        errors.append(
                            f"{business_id}/"
                            f"{edge}/"
                            f"{token_key}: "
                            f"{clean}"
                        )

                if not any_success:
                    print(
                        "WARNING: all tokens "
                        "failed for "
                        f"{business_id}/"
                        f"{edge}"
                    )

        accounts = sorted(
            collected.values(),
            key=lambda item: (
                item[
                    "account_name"
                ].casefold()
            ),
        )

        if not accounts:
            raise RuntimeError(
                "No Meta Ad Accounts "
                "discovered.\n"
                + "\n".join(
                    errors[-10:]
                )
            )

        print(
            "Meta discovery: "
            f"{len(accounts)} "
            "unique Ad Accounts"
        )

        return accounts


    def fetch_today_spend(
        self,
        account: dict[str, Any],
        today: str,
    ) -> float:

        account_id = (
            normalize_account_id(
                account["id"]
            )
        )

        url = (
            f"{BASE_URL}/"
            f"{self.api_version}/"
            f"act_{account_id}/"
            f"insights"
        )

        params = {
            "fields": "spend",
            "level": "account",
            "time_range": json.dumps(
                {
                    "since": today,
                    "until": today,
                },
                separators=(
                    ",",
                    ":",
                ),
            ),
            "limit": 100,
        }

        preferred = str(
            account.get(
                "token_key"
            )
            or "token_1"
        )

        keys = [
            preferred
        ] + [
            key
            for key in self.tokens
            if key != preferred
        ]

        failures: list[str] = []

        for key in keys:

            token = self.tokens.get(
                key
            )

            if not token:
                continue

            try:

                rows = fetch_all_pages(
                    url,
                    {
                        **params,
                        "access_token": token,
                    },
                )

                return sum(
                    to_float(
                        row.get(
                            "spend"
                        )
                    )
                    for row in rows
                )

            except Exception as exc:

                clean = str(exc).replace(
                    token,
                    "[REDACTED]",
                )

                failures.append(
                    f"{key}: {clean}"
                )

        raise RuntimeError(
            f"act_{account_id}: "
            + " | ".join(
                failures
            )
        )


    def fetch_all_today_spend(
        self,
        accounts: list[
            dict[str, Any]
        ],
        today: str,
    ) -> tuple[
        list[dict[str, Any]],
        list[str],
    ]:

        output: list[
            dict[str, Any]
        ] = []

        errors: list[str] = []

        with ThreadPoolExecutor(
            max_workers=min(
                self.max_workers,
                max(
                    1,
                    len(accounts),
                ),
            )
        ) as executor:

            future_map = {
                executor.submit(
                    self.fetch_today_spend,
                    account,
                    today,
                ): account
                for account
                in accounts
            }

            for future in as_completed(
                future_map
            ):

                account = (
                    future_map[
                        future
                    ]
                )

                try:

                    spend = (
                        future.result()
                    )

                    output.append({
                        **account,
                        "spend": spend,
                    })

                except Exception as exc:

                    errors.append(
                        f"{account.get('account_name')}: "
                        f"{exc}"
                    )

        return output, errors


# ============================================================
# GOOGLE SHEET CLIENT
# ============================================================

class TrackerSheet:

    def __init__(self) -> None:

        raw_json = required_env(
            "GOOGLE_SERVICE_ACCOUNT_JSON"
        )

        try:
            service_info = json.loads(
                raw_json
            )

        except json.JSONDecodeError as exc:

            raise RuntimeError(
                "GOOGLE_SERVICE_ACCOUNT_JSON "
                "is not valid JSON."
            ) from exc

        credentials = (
            Credentials
            .from_service_account_info(
                service_info,
                scopes=GOOGLE_SCOPES,
            )
        )

        gc = gspread.authorize(
            credentials
        )

        self.book = gc.open_by_key(
            required_env(
                "GOOGLE_SHEET_ID"
            )
        )

        try:
            self.ws = self.book.worksheet(
                SHEET_TAB
            )

        except gspread.WorksheetNotFound:

            raise RuntimeError(
                f"Tab '{SHEET_TAB}' "
                "was not found. "
                "Upload the provided "
                "template first."
            )


    def batch_write(
        self,
        data: list[
            dict[str, Any]
        ],
        *,
        value_input_option: str = "USER_ENTERED",
        attempts: int = 5,
    ) -> None:
        """
        Write many ranges in ONE Google Sheets request.

        This is the main protection against Sheets write quota errors.

        Example data item:
            {
                "range": "A1:B2",
                "values": [[1, 2], [3, 4]]
            }

        On HTTP 429, wait and retry automatically.
        """

        if not data:
            return

        last_error: Exception | None = None

        for attempt in range(
            1,
            attempts + 1,
        ):
            try:
                self.ws.batch_update(
                    data,
                    value_input_option=(
                        value_input_option
                    ),
                )
                return

            except gspread.exceptions.APIError as exc:
                last_error = exc

                response = getattr(
                    exc,
                    "response",
                    None,
                )

                status_code = getattr(
                    response,
                    "status_code",
                    None,
                )

                is_quota_error = (
                    status_code == 429
                    or "429" in str(exc)
                    or "Quota exceeded"
                    in str(exc)
                )

                if not is_quota_error:
                    raise

                if attempt >= attempts:
                    break

                wait_seconds = min(
                    60,
                    10 * (
                        2 ** (
                            attempt - 1
                        )
                    ),
                )

                print(
                    "Google Sheets write quota "
                    f"hit (429). Waiting "
                    f"{wait_seconds}s before "
                    f"retry {attempt + 1}/"
                    f"{attempts}..."
                )

                time.sleep(
                    wait_seconds
                )

        raise RuntimeError(
            "Google Sheets write quota "
            "still exceeded after retries: "
            f"{last_error}"
        )


    # --------------------------------------------------------
    # AGENT DEFINITIONS
    # --------------------------------------------------------

    def read_agent_definitions(
        self,
    ) -> list[dict[str, str]]:

        headers = (
            self.ws.get(
                INPUT_HEADER_RANGE
            )
            or [[]]
        )[0]

        headers += [
            ""
        ] * (
            13
            - len(headers)
        )

        agents: list[
            dict[str, str]
        ] = []

        # C:K = index 2..10
        for idx in range(
            2,
            11,
        ):

            parsed = (
                parse_agent_header(
                    headers[idx]
                )
            )

            if not parsed:
                continue

            code, name = parsed

            agents.append({
                "code": code,
                "name": name,
                "header": str(
                    headers[idx]
                ),
            })

        if not agents:

            raise RuntimeError(
                "No Agents found "
                "in C8:K8. "
                "Use headers like "
                "AA - Abdallah."
            )

        return agents


    # --------------------------------------------------------
    # INTERNAL HEADERS
    # --------------------------------------------------------

    def sync_internal_headers(
        self,
        agents: list[
            dict[str, str]
        ],
    ) -> None:

        history_headers = [
            "Date",
            "Overall Budget",
            *[
                agent["header"]
                for agent
                in agents
            ],
            "Notes",
            "Saved At",
        ]

        raw_headers = [
            "Date",
            "Hour",
            "Scope",
            "Agent Code",
            "Agent Name",
            "Spend",
            "Budget",
            "Spend %",
            "Remaining",
            "Updated At",
        ]

        self.batch_write([
            {
                "range": (
                    HISTORY_HEADER_RANGE
                ),
                "values": [
                    history_headers
                ],
            },
            {
                "range": (
                    RAW_HEADER_RANGE
                ),
                "values": [
                    raw_headers
                ],
            },
        ])


    # --------------------------------------------------------
    # PROFESSIONAL DESIGN REPAIR — ONE TIME ONLY
    # --------------------------------------------------------

    def apply_professional_design(
        self,
        agents: list[
            dict[str, str]
        ],
    ) -> None:
        """
        Fix all visual leftovers from previous layouts in one Sheets
        batchUpdate request:

        - Remove old merged cells that collide with new tables.
        - Remove old conditional-format rules that were coloring Spend.
        - Remove the old Agent dropdown from the Spend hour header.
        - Merge the new section titles correctly.
        - Apply full-width styles to every header.
        - Apply correct number formats.
        - Add clean percentage heatmaps only to percentage tables.
        - Add proper Agent selector dropdown at B83.
        - Hide internal Z:AX helper columns.

        This runs only when LAYOUT_VERSION changes, not every hour.
        """

        sheet_id = self.ws.id

        # ----------------------------------------------------
        # Read current merges + conditional rules
        # ----------------------------------------------------
        metadata = self.book.fetch_sheet_metadata(
            params={
                "fields": (
                    "sheets("
                    "properties(sheetId),"
                    "merges,"
                    "conditionalFormats"
                    ")"
                )
            }
        )

        current_sheet_meta: dict[str, Any] = {}

        for item in metadata.get(
            "sheets",
            [],
        ):
            if (
                item.get(
                    "properties",
                    {},
                ).get(
                    "sheetId"
                )
                == sheet_id
            ):
                current_sheet_meta = item
                break

        requests_list: list[
            dict[str, Any]
        ] = []

        # ----------------------------------------------------
        # Delete old conditional formatting.
        #
        # We recreate the two correct heatmaps below.
        # Delete from last index -> first because indexes shift.
        # ----------------------------------------------------
        conditional_rules = (
            current_sheet_meta.get(
                "conditionalFormats",
                [],
            )
            or []
        )

        for index in reversed(
            range(
                len(
                    conditional_rules
                )
            )
        ):
            requests_list.append({
                "deleteConditionalFormatRule": {
                    "sheetId": (
                        sheet_id
                    ),
                    "index": index,
                }
            })

        # ----------------------------------------------------
        # Remove existing merges in the redesigned lower area.
        #
        # This is important because the OLD Agent selector was
        # B48:E48. B48:E48 is now our Spend hour header and was
        # causing 12:55 / 1:55 / 2:55 / 3:55 to collapse.
        # ----------------------------------------------------
        for merge_range in (
            current_sheet_meta.get(
                "merges",
                [],
            )
            or []
        ):
            start_row = (
                merge_range.get(
                    "startRowIndex",
                    0,
                )
            )
            end_row = (
                merge_range.get(
                    "endRowIndex",
                    start_row + 1,
                )
            )

            start_col = (
                merge_range.get(
                    "startColumnIndex",
                    0,
                )
            )
            end_col = (
                merge_range.get(
                    "endColumnIndex",
                    start_col + 1,
                )
            )

            # Redesigned visual zone = rows 47:129, columns A:Y
            intersects = (
                end_row > 46
                and start_row < 129
                and end_col > 0
                and start_col < 25
            )

            if intersects:
                requests_list.append({
                    "unmergeCells": {
                        "range": (
                            merge_range
                        )
                    }
                })

        # ----------------------------------------------------
        # Clear the old dropdown/data validation from B48:E48.
        # ----------------------------------------------------
        requests_list.append({
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 47,
                    "endRowIndex": 48,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5,
                },
                "rule": None,
            }
        })

        # ----------------------------------------------------
        # Reset visual formatting in redesigned area.
        # This wipes old purple/green/orange formatting residue.
        # ----------------------------------------------------
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 46,
                    "endRowIndex": 129,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1,
                            "green": 1,
                            "blue": 1,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 0.08,
                                "green": 0.10,
                                "blue": 0.15,
                            },
                            "fontSize": 9,
                            "bold": False,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                        "wrapStrategy": "CLIP",
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "backgroundColor,"
                    "userEnteredFormat."
                    "textFormat,"
                    "userEnteredFormat."
                    "horizontalAlignment,"
                    "userEnteredFormat."
                    "verticalAlignment,"
                    "userEnteredFormat."
                    "wrapStrategy"
                ),
            }
        })

        # ----------------------------------------------------
        # Also standardize existing Overall Percentage table
        # header and number format.
        # ----------------------------------------------------
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 12,
                    "endRowIndex": 13,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.10,
                            "green": 0.38,
                            "blue": 0.88,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 9,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "backgroundColor,"
                    "userEnteredFormat."
                    "textFormat,"
                    "userEnteredFormat."
                    "horizontalAlignment,"
                    "userEnteredFormat."
                    "verticalAlignment"
                ),
            }
        })

        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 13,
                    "endRowIndex": 44,
                    "startColumnIndex": 1,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": (
                                "PERCENT"
                            ),
                            "pattern": (
                                "0.00%"
                            ),
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "numberFormat,"
                    "userEnteredFormat."
                    "horizontalAlignment"
                ),
            }
        })

        # ----------------------------------------------------
        # SPEND TITLE A47:Y47
        # ----------------------------------------------------
        requests_list.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 46,
                    "endRowIndex": 47,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "mergeType": "MERGE_ALL",
            }
        })

        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 46,
                    "endRowIndex": 47,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.07,
                            "green": 0.10,
                            "blue": 0.16,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 11,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Spend hour header A48:Y48
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 47,
                    "endRowIndex": 48,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.10,
                            "green": 0.38,
                            "blue": 0.88,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 9,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                        "wrapStrategy": "CLIP",
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Spend data number format B49:Y79
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 48,
                    "endRowIndex": 79,
                    "startColumnIndex": 1,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": (
                                "#,##0.00"
                            ),
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "numberFormat,"
                    "userEnteredFormat."
                    "horizontalAlignment"
                ),
            }
        })

        # Day labels for Spend
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 48,
                    "endRowIndex": 79,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.96,
                            "green": 0.97,
                            "blue": 0.98,
                        },
                        "textFormat": {
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # ----------------------------------------------------
        # AGENT TITLE A82:Y82
        # ----------------------------------------------------
        requests_list.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 81,
                    "endRowIndex": 82,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "mergeType": "MERGE_ALL",
            }
        })

        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 81,
                    "endRowIndex": 82,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.30,
                            "green": 0.12,
                            "blue": 0.62,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 11,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Agent selector label A83
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 82,
                    "endRowIndex": 83,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.20,
                            "green": 0.22,
                            "blue": 0.27,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Merge B83:E83 for a clean selector box.
        requests_list.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 82,
                    "endRowIndex": 83,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5,
                },
                "mergeType": "MERGE_ALL",
            }
        })

        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 82,
                    "endRowIndex": 83,
                    "startColumnIndex": 1,
                    "endColumnIndex": 5,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.95,
                            "green": 0.93,
                            "blue": 1,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 0.16,
                                "green": 0.08,
                                "blue": 0.35,
                            },
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Correct dropdown at B83
        validation_values = [
            {
                "userEnteredValue": (
                    agent[
                        "header"
                    ]
                )
            }
            for agent in agents
        ]

        requests_list.append({
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 82,
                    "endRowIndex": 83,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2,
                },
                "rule": {
                    "condition": {
                        "type": (
                            "ONE_OF_LIST"
                        ),
                        "values": (
                            validation_values
                        ),
                    },
                    "strict": True,
                    "showCustomUi": True,
                },
            }
        })

        # Agent hour header A85:Y85 — FULL WIDTH, not only A:I
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 84,
                    "endRowIndex": 85,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.43,
                            "green": 0.14,
                            "blue": 0.78,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 9,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Agent percentage data
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 85,
                    "endRowIndex": 116,
                    "startColumnIndex": 1,
                    "endColumnIndex": 25,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": (
                                "PERCENT"
                            ),
                            "pattern": (
                                "0.00%"
                            ),
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "numberFormat,"
                    "userEnteredFormat."
                    "horizontalAlignment"
                ),
            }
        })

        # Agent day labels
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 85,
                    "endRowIndex": 116,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.96,
                            "green": 0.97,
                            "blue": 0.98,
                        },
                        "textFormat": {
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # ----------------------------------------------------
        # SNAPSHOT TITLE A119:I119
        # ----------------------------------------------------
        requests_list.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 118,
                    "endRowIndex": 119,
                    "startColumnIndex": 0,
                    "endColumnIndex": 9,
                },
                "mergeType": "MERGE_ALL",
            }
        })

        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 118,
                    "endRowIndex": 119,
                    "startColumnIndex": 0,
                    "endColumnIndex": 9,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.05,
                            "green": 0.33,
                            "blue": 0.30,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 11,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Snapshot headers A120:I120
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 119,
                    "endRowIndex": 120,
                    "startColumnIndex": 0,
                    "endColumnIndex": 9,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.07,
                            "green": 0.47,
                            "blue": 0.42,
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                            },
                            "fontSize": 9,
                            "bold": True,
                        },
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                    }
                },
                "fields": (
                    "userEnteredFormat"
                ),
            }
        })

        # Snapshot data row alignment/wrapping
        requests_list.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 120,
                    "endRowIndex": 129,
                    "startColumnIndex": 0,
                    "endColumnIndex": 9,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": (
                            "CENTER"
                        ),
                        "verticalAlignment": (
                            "MIDDLE"
                        ),
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": (
                    "userEnteredFormat."
                    "horizontalAlignment,"
                    "userEnteredFormat."
                    "verticalAlignment,"
                    "userEnteredFormat."
                    "wrapStrategy"
                ),
            }
        })

        # Snapshot number formats
        requests_list.extend([
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 120,
                        "endRowIndex": 129,
                        "startColumnIndex": 2,
                        "endColumnIndex": 4,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": (
                                    "NUMBER"
                                ),
                                "pattern": (
                                    "#,##0.00"
                                ),
                            }
                        }
                    },
                    "fields": (
                        "userEnteredFormat."
                        "numberFormat"
                    ),
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 120,
                        "endRowIndex": 129,
                        "startColumnIndex": 4,
                        "endColumnIndex": 5,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": (
                                    "PERCENT"
                                ),
                                "pattern": (
                                    "0.00%"
                                ),
                            }
                        }
                    },
                    "fields": (
                        "userEnteredFormat."
                        "numberFormat"
                    ),
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 120,
                        "endRowIndex": 129,
                        "startColumnIndex": 5,
                        "endColumnIndex": 6,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": (
                                    "NUMBER"
                                ),
                                "pattern": (
                                    "#,##0.00"
                                ),
                            }
                        }
                    },
                    "fields": (
                        "userEnteredFormat."
                        "numberFormat"
                    ),
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 120,
                        "endRowIndex": 129,
                        "startColumnIndex": 6,
                        "endColumnIndex": 8,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": (
                                    "PERCENT"
                                ),
                                "pattern": (
                                    "0.00%"
                                ),
                            }
                        }
                    },
                    "fields": (
                        "userEnteredFormat."
                        "numberFormat"
                    ),
                }
            },
        ])

        # ----------------------------------------------------
        # Borders around visible data tables
        # ----------------------------------------------------
        border_style = {
            "style": "SOLID",
            "color": {
                "red": 0.84,
                "green": 0.86,
                "blue": 0.90,
            },
        }

        for grid_range in (
            {
                "sheetId": sheet_id,
                "startRowIndex": 12,
                "endRowIndex": 44,
                "startColumnIndex": 0,
                "endColumnIndex": 25,
            },
            {
                "sheetId": sheet_id,
                "startRowIndex": 47,
                "endRowIndex": 79,
                "startColumnIndex": 0,
                "endColumnIndex": 25,
            },
            {
                "sheetId": sheet_id,
                "startRowIndex": 84,
                "endRowIndex": 116,
                "startColumnIndex": 0,
                "endColumnIndex": 25,
            },
            {
                "sheetId": sheet_id,
                "startRowIndex": 119,
                "endRowIndex": 129,
                "startColumnIndex": 0,
                "endColumnIndex": 9,
            },
        ):
            requests_list.append({
                "updateBorders": {
                    "range": grid_range,
                    "top": border_style,
                    "bottom": border_style,
                    "left": border_style,
                    "right": border_style,
                    "innerHorizontal": (
                        border_style
                    ),
                    "innerVertical": (
                        border_style
                    ),
                }
            })

        # ----------------------------------------------------
        # Column widths:
        # A = Day / labels
        # B:Y = all 24 hour columns consistently.
        # ----------------------------------------------------
        requests_list.extend([
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": (
                            "COLUMNS"
                        ),
                        "startIndex": 0,
                        "endIndex": 1,
                    },
                    "properties": {
                        "pixelSize": 64,
                    },
                    "fields": (
                        "pixelSize"
                    ),
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": (
                            "COLUMNS"
                        ),
                        "startIndex": 1,
                        "endIndex": 25,
                    },
                    "properties": {
                        "pixelSize": 82,
                    },
                    "fields": (
                        "pixelSize"
                    ),
                }
            },
        ])

        # Title/header row heights.
        for start_index, end_index, size in (
            (46, 47, 28),
            (47, 48, 24),
            (81, 82, 28),
            (82, 83, 26),
            (84, 85, 24),
            (118, 119, 28),
            (119, 120, 25),
        ):
            requests_list.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": (
                            start_index
                        ),
                        "endIndex": (
                            end_index
                        ),
                    },
                    "properties": {
                        "pixelSize": size,
                    },
                    "fields": "pixelSize",
                }
            })

        # ----------------------------------------------------
        # Hide internal helper columns Z:AX.
        # The user-facing dashboard remains A:Y only.
        # ----------------------------------------------------
        requests_list.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 25,
                    "endIndex": 50,
                },
                "properties": {
                    "hiddenByUser": True,
                },
                "fields": "hiddenByUser",
            }
        })

        # ----------------------------------------------------
        # Recreate CLEAN percentage heatmaps only.
        # No heatmap on Spend table.
        # ----------------------------------------------------
        percentage_ranges = [
            {
                "sheetId": sheet_id,
                "startRowIndex": 13,
                "endRowIndex": 44,
                "startColumnIndex": 1,
                "endColumnIndex": 25,
            },
            {
                "sheetId": sheet_id,
                "startRowIndex": 85,
                "endRowIndex": 116,
                "startColumnIndex": 1,
                "endColumnIndex": 25,
            },
        ]

        for grid_range in (
            percentage_ranges
        ):
            requests_list.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [
                            grid_range
                        ],
                        "gradientRule": {
                            "minpoint": {
                                "color": {
                                    "red": 0.86,
                                    "green": 0.93,
                                    "blue": 1.00,
                                },
                                "type": "NUMBER",
                                "value": "0",
                            },
                            "midpoint": {
                                "color": {
                                    "red": 1.00,
                                    "green": 0.94,
                                    "blue": 0.70,
                                },
                                "type": "NUMBER",
                                "value": "0.5",
                            },
                            "maxpoint": {
                                "color": {
                                    "red": 1.00,
                                    "green": 0.80,
                                    "blue": 0.80,
                                },
                                "type": "NUMBER",
                                "value": "1",
                            },
                        },
                    },
                    "index": 0,
                }
            })

        # ----------------------------------------------------
        # One Google Sheets batchUpdate for the full visual repair.
        # ----------------------------------------------------
        if requests_list:
            self.book.batch_update({
                "requests": (
                    requests_list
                )
            })


    # --------------------------------------------------------
    # VISIBLE LAYOUT / ONE-TIME MIGRATION
    # --------------------------------------------------------

    def ensure_visible_layout(
        self,
        agents: list[
            dict[str, str]
        ],
    ) -> None:
        """
        Migrate/repair the visible dashboard exactly once for V6.

        Hourly runs after migration do NOT repeat this formatting.
        """

        current_version = str(
            self.ws.acell(
                LAYOUT_VERSION_CELL
            ).value
            or ""
        ).strip()

        if (
            current_version
            == LAYOUT_VERSION
        ):
            return

        # Prefer the current V5 selector at B83.
        # Fall back to old B48 only for older layouts.
        current_selector = str(
            self.ws.acell(
                AGENT_SELECTOR_CELL
            ).value
            or ""
        ).strip()

        if not current_selector:
            current_selector = str(
                self.ws.acell(
                    "B48"
                ).value
                or ""
            ).strip()

        parsed_selector = (
            parse_agent_header(
                current_selector
            )
        )

        valid_codes = {
            agent["code"]
            for agent in agents
        }

        if (
            parsed_selector
            and parsed_selector[0]
            in valid_codes
        ):
            selected_header = (
                current_selector
            )
        else:
            selected_header = (
                agents[0][
                    "header"
                ]
            )

        # Values in the matrices can safely be rebuilt from the Raw Log.
        # Clear the redesigned lower visible area before rebuilding.
        try:
            self.ws.batch_clear([
                "A47:Y129",
            ])
        except Exception:
            pass

        day_values = [
            [day]
            for day in range(
                1,
                32,
            )
        ]

        # Update the carry-forward budget title too.
        self.batch_write([
            {
                "range": "A7",
                "values": [[
                    "BUDGET / ALLOCATION CHANGE "
                    "— EDIT ONLY WHEN VALUES CHANGE"
                ]],
            },
            {
                "range": (
                    SPEND_TITLE_CELL
                ),
                "values": [[
                    "LIVE OVERALL SPEND "
                    "— DAY × HOUR"
                ]],
            },
            {
                "range": (
                    AGENT_TITLE_CELL
                ),
                "values": [[
                    "AGENT PACING VIEW "
                    "— SELECT ONE AGENT"
                ]],
            },
            {
                "range": (
                    AGENT_SELECTOR_LABEL_CELL
                ),
                "values": [[
                    "Agent:"
                ]],
            },
            {
                "range": (
                    AGENT_SELECTOR_CELL
                ),
                "values": [[
                    selected_header
                ]],
            },
            {
                "range": (
                    SNAPSHOT_TITLE_CELL
                ),
                "values": [[
                    "LIVE AGENT SNAPSHOT"
                ]],
            },
            {
                "range": (
                    SNAPSHOT_HEADER_RANGE
                ),
                "values": [[
                    "Code",
                    "Agent",
                    "Budget",
                    "Spend",
                    "Used %",
                    "Remaining",
                    "Expected Pace %",
                    "Pace Delta",
                    "Status",
                ]],
            },
            {
                "range": "A14:A44",
                "values": day_values,
            },
            {
                "range": "A49:A79",
                "values": day_values,
            },
            {
                "range": "A86:A116",
                "values": day_values,
            },
        ])

        # One-time structural/design repair.
        self.apply_professional_design(
            agents
        )

        # Mark migration only AFTER successful design repair.
        self.batch_write([
            {
                "range": (
                    LAYOUT_VERSION_CELL
                ),
                "values": [[
                    LAYOUT_VERSION
                ]],
            },
        ])


    # --------------------------------------------------------
    # AUTOMATED HOUR HEADERS
    # --------------------------------------------------------

    def sync_hour_headers(
        self,
        now: datetime,
    ) -> None:
        """
        Maintain the 24 hourly columns automatically in 12-hour format.

        Example with cron around :55:
            12:55 AM, 1:55 AM ... 12:55 PM ... 11:55 PM

        All three visible tables receive the same header row in ONE write.
        """

        existing = (
            self.ws.get(
                OVERALL_HOUR_HEADER_RANGE
            )
            or [[]]
        )[0]

        existing += [
            ""
        ] * (
            24
            - len(existing)
        )

        headers: list[str] = []

        for hour in range(24):
            minute = header_minute(
                existing[hour],
                default=0,
            )

            headers.append(
                format_hour_12h(
                    hour,
                    minute,
                )
            )

        headers[
            now.hour
        ] = format_hour_12h(
            now.hour,
            now.minute,
        )

        self.batch_write([
            {
                "range": (
                    OVERALL_DAY_HEADER_CELL
                ),
                "values": [["Day"]],
            },
            {
                "range": (
                    OVERALL_HOUR_HEADER_RANGE
                ),
                "values": [headers],
            },
            {
                "range": (
                    SPEND_DAY_HEADER_CELL
                ),
                "values": [["Day"]],
            },
            {
                "range": (
                    SPEND_HOUR_HEADER_RANGE
                ),
                "values": [headers],
            },
            {
                "range": (
                    AGENT_DAY_HEADER_CELL
                ),
                "values": [["Day"]],
            },
            {
                "range": (
                    AGENT_HOUR_HEADER_RANGE
                ),
                "values": [headers],
            },
        ])


    # --------------------------------------------------------
    # SAVE DAILY INPUT TO HISTORY
    # --------------------------------------------------------

    def save_input_to_history(
        self,
        agents: list[
            dict[str, str]
        ],
        saved_at: str,
    ) -> str:

        row = (
            self.ws.get(
                INPUT_ROW_RANGE
            )
            or [[]]
        )[0]

        row += [
            ""
        ] * (
            13
            - len(row)
        )

        budget_date = (
            parse_sheet_date(
                row[0]
            )
        )

        if not budget_date:
            self.batch_write([
                {
                    "range": (
                        INPUT_STATUS_CELL
                    ),
                    "values": [[
                        "NO DATE"
                    ]],
                },
            ])
            return ""

        meaningful_values = [
            str(
                value or ""
            ).strip()
            for value
            in row[1:11]
        ]

        if not any(
            meaningful_values
        ):
            self.batch_write([
                {
                    "range": (
                        INPUT_STATUS_CELL
                    ),
                    "values": [[
                        "NO BUDGETS ENTERED"
                    ]],
                },
            ])
            return budget_date

        history_values = (
            self.ws.get(
                f"AA"
                f"{HISTORY_DATA_START}"
                f":AM"
            )
        )

        target_row: int | None = (
            None
        )

        for (
            offset,
            existing,
        ) in enumerate(
            history_values,
            start=HISTORY_DATA_START,
        ):
            if not existing:
                continue

            existing_date = (
                parse_sheet_date(
                    existing[0]
                )
            )

            if (
                existing_date
                == budget_date
            ):
                target_row = offset

        output = [
            budget_date,
            row[1],
            *row[2:11],
            row[11],
            saved_at,
        ]

        if target_row is not None:
            existing_padded = (
                list(
                    history_values[
                        target_row
                        - HISTORY_DATA_START
                    ]
                )
                + [""] * 13
            )

            old_values = [
                str(
                    value or ""
                ).strip()
                for value
                in existing_padded[
                    :12
                ]
            ]

            new_values = [
                str(
                    value or ""
                ).strip()
                for value
                in output[
                    :12
                ]
            ]

            if (
                old_values
                == new_values
            ):
                self.batch_write([
                    {
                        "range": (
                            INPUT_STATUS_CELL
                        ),
                        "values": [[
                            f"ACTIVE FROM "
                            f"{budget_date}"
                        ]],
                    },
                ])

                return budget_date

        if target_row is None:
            target_row = (
                HISTORY_DATA_START
                + len(
                    history_values
                )
            )

        self.batch_write([
            {
                "range": (
                    f"AA{target_row}"
                    f":AM{target_row}"
                ),
                "values": [output],
            },
            {
                "range": (
                    INPUT_STATUS_CELL
                ),
                "values": [[
                    f"SAVED CHANGE "
                    f"{budget_date}"
                ]],
            },
        ])

        return budget_date


    # --------------------------------------------------------
    # READ BUDGET FOR ANY DATE
    # --------------------------------------------------------

    def read_budget_for_date(
        self,
        target_date: str,
        agents: list[
            dict[str, str]
        ],
    ) -> dict[str, Any]:
        """
        Carry-forward budget logic.

        The budget does NOT need to be entered every day.

        Example:
            2026-10-02 = 90,000
            2026-10-03 = no new row
            2026-10-04 = no new row

        Effective budget on Oct 3 and Oct 4 remains 90,000.

        When a newer dated budget is saved, it becomes the new default
        from that date onward until another change is saved.
        """

        rows = self.ws.get(
            f"AA"
            f"{HISTORY_DATA_START}"
            f":AM"
        )

        target_dt = datetime.strptime(
            target_date,
            "%Y-%m-%d",
        ).date()

        latest_date: str | None = None
        latest_row: list[Any] | None = None

        for row in rows:
            padded = (
                list(row)
                + [""] * (
                    13
                    - len(row)
                )
            )

            row_date = (
                parse_sheet_date(
                    padded[0]
                )
            )

            if not re.fullmatch(
                r"\d{4}-\d{2}-\d{2}",
                row_date,
            ):
                continue

            try:
                row_dt = datetime.strptime(
                    row_date,
                    "%Y-%m-%d",
                ).date()
            except ValueError:
                continue

            # Only budgets effective on or before the requested date.
            if row_dt > target_dt:
                continue

            if (
                latest_date is None
                or row_date > latest_date
            ):
                latest_date = row_date
                latest_row = padded

        allocations: dict[
            str,
            float | None,
        ] = {
            agent["code"]: None
            for agent in agents
        }

        overall: float | None = None

        if latest_row is not None:

            overall_raw = str(
                latest_row[1]
                or ""
            ).strip()

            if overall_raw:
                candidate = to_float(
                    overall_raw,
                    default=float("nan"),
                )

                if (
                    math.isfinite(candidate)
                    and candidate >= 0
                ):
                    overall = candidate

            for (
                idx,
                agent,
            ) in enumerate(
                agents,
                start=2,
            ):

                raw_value = str(
                    latest_row[idx]
                    or ""
                ).strip()

                if not raw_value:
                    continue

                candidate = to_float(
                    raw_value,
                    default=float("nan"),
                )

                if (
                    math.isfinite(candidate)
                    and candidate >= 0
                ):
                    allocations[
                        agent["code"]
                    ] = candidate

        valid_agent_budgets = [
            value
            for value in allocations.values()
            if value is not None
        ]

        if (
            overall is None
            and valid_agent_budgets
        ):
            overall = sum(
                valid_agent_budgets
            )

        return {
            "overall": overall,
            "agents": allocations,
            "found": (
                latest_row
                is not None
            ),
            "source_date": latest_date,
            "inherited": (
                latest_date is not None
                and latest_date != target_date
            ),
        }


    def read_budget_history(
        self,
        agents: list[
            dict[str, str]
        ],
    ) -> dict[
        str,
        dict[str, Any],
    ]:
        """
        Load budget CHANGE POINTS.

        Each dated row means:
        "Use these budgets from this date onward until another dated
        budget row replaces them."

        So there is no need to add one row every day.
        """

        rows = self.ws.get(
            f"AA"
            f"{HISTORY_DATA_START}"
            f":AM"
        )

        result: dict[
            str,
            dict[str, Any],
        ] = {}

        for row in rows:
            padded = (
                list(row)
                + [""] * (
                    13
                    - len(row)
                )
            )

            date_value = (
                parse_sheet_date(
                    padded[0]
                )
            )

            if not re.fullmatch(
                r"\d{4}-\d{2}-\d{2}",
                date_value,
            ):
                continue

            overall: float | None = None

            overall_raw = str(
                padded[1]
                or ""
            ).strip()

            if overall_raw:
                candidate = to_float(
                    overall_raw,
                    default=float("nan"),
                )

                if (
                    math.isfinite(candidate)
                    and candidate >= 0
                ):
                    overall = candidate

            allocations: dict[
                str,
                float | None,
            ] = {
                agent["code"]: None
                for agent in agents
            }

            for idx, agent in enumerate(
                agents,
                start=2,
            ):
                raw_value = str(
                    padded[idx]
                    or ""
                ).strip()

                if not raw_value:
                    continue

                candidate = to_float(
                    raw_value,
                    default=float("nan"),
                )

                if (
                    math.isfinite(candidate)
                    and candidate >= 0
                ):
                    allocations[
                        agent["code"]
                    ] = candidate

            valid_agent_budgets = [
                value
                for value in allocations.values()
                if value is not None
            ]

            if (
                overall is None
                and valid_agent_budgets
            ):
                overall = sum(
                    valid_agent_budgets
                )

            result[
                date_value
            ] = {
                "overall": overall,
                "agents": allocations,
            }

        return result


    @staticmethod
    def effective_budget_from_history(
        budget_history: dict[
            str,
            dict[str, Any],
        ],
        target_date: str,
        *,
        scope: str,
        agent_code: str | None = None,
    ) -> float | None:
        """
        Resolve the latest budget change whose date is <= target_date.
        """

        effective_date: str | None = None
        effective_record: dict[str, Any] | None = None

        for date_value, record in budget_history.items():

            if date_value > target_date:
                continue

            if (
                effective_date is None
                or date_value > effective_date
            ):
                effective_date = date_value
                effective_record = record

        if effective_record is None:
            return None

        if scope.upper() == "OVERALL":
            return effective_record.get(
                "overall"
            )

        if agent_code is None:
            return None

        return (
            effective_record
            .get(
                "agents",
                {},
            )
            .get(
                agent_code
            )
        )


    # --------------------------------------------------------
    # RAW HOURLY LOG
    # --------------------------------------------------------

    def upsert_raw_rows(
        self,
        rows: list[
            list[Any]
        ],
        today: str,
        hour_bucket: str,
    ) -> None:
        """
        Upsert ALL Overall + Agent raw rows in ONE Sheets write request.

        Unique key:
            Date + Hour Bucket + Agent Code
        """

        existing = self.ws.get(
            f"AO"
            f"{RAW_DATA_START}"
            f":AX"
        )

        index_by_key: dict[
            tuple[
                str,
                str,
                str,
            ],
            int,
        ] = {}

        for (
            offset,
            existing_row,
        ) in enumerate(
            existing,
            start=RAW_DATA_START,
        ):

            padded = (
                list(
                    existing_row
                )
                + [""] * 10
            )

            key = (
                parse_sheet_date(
                    padded[0]
                ),
                str(
                    padded[1]
                    or ""
                ).strip(),
                str(
                    padded[3]
                    or ""
                )
                .strip()
                .upper(),
            )

            if all(key):
                index_by_key[
                    key
                ] = offset

        batch_data: list[
            dict[str, Any]
        ] = []

        append_rows: list[
            list[Any]
        ] = []

        for row in rows:

            code = str(
                row[3]
            ).strip().upper()

            key = (
                today,
                hour_bucket,
                code,
            )

            target = (
                index_by_key
                .get(key)
            )

            if target:
                batch_data.append({
                    "range": (
                        f"AO{target}"
                        f":AX{target}"
                    ),
                    "values": [row],
                })
            else:
                append_rows.append(
                    row
                )

        if append_rows:
            next_row = (
                RAW_DATA_START
                + len(existing)
            )

            end_row = (
                next_row
                + len(
                    append_rows
                )
                - 1
            )

            batch_data.append({
                "range": (
                    f"AO{next_row}"
                    f":AX{end_row}"
                ),
                "values": (
                    append_rows
                ),
            })

        self.batch_write(
            batch_data
        )


    def raw_rows(
        self,
    ) -> list[list[Any]]:
        """
        Read FORMATTED values from Google Sheets.

        This keeps dates as strings like 2026-10-01 instead of Google
        spreadsheet serial numbers. Percentages may come back like 66.97%,
        and to_float() safely converts them to 0.6697.
        """

        return self.ws.get(
            f"AO"
            f"{RAW_DATA_START}"
            f":AX",
            value_render_option=(
                "FORMATTED_VALUE"
            ),
        )


    # --------------------------------------------------------
    # MATRIX BUILDING
    # --------------------------------------------------------

    @staticmethod
    def build_matrix(
        raw_rows: list[
            list[Any]
        ],
        year: int,
        month: int,
        *,
        scope: str,
        metric: str,
        budget_history: dict[
            str,
            dict[str, Any],
        ],
        agent_code: (
            str | None
        ) = None,
    ) -> list[list[Any]]:
        """
        metric:
            "percent" -> Spend / the saved budget for THAT DATE
            "spend"   -> cumulative Spend Today

        Percentages are rebuilt from raw SPEND + Budget History whenever
        possible. This means if the budget is entered later, older hourly
        snapshots automatically backfill on the next run.
        """

        matrix: list[
            list[Any]
        ] = [
            [""] * 24
            for _ in range(31)
        ]

        for row in raw_rows:
            padded = (
                list(row)
                + [""] * (
                    10
                    - len(row)
                )
            )

            date_value = (
                parse_sheet_date(
                    padded[0]
                )
            )

            if not re.fullmatch(
                r"\d{4}-\d{2}-\d{2}",
                date_value,
            ):
                continue

            try:
                dt = datetime.strptime(
                    date_value,
                    "%Y-%m-%d",
                )
            except ValueError:
                continue

            if (
                dt.year != year
                or dt.month != month
            ):
                continue

            row_scope = str(
                padded[2]
                or ""
            ).strip().upper()

            if (
                row_scope
                != scope.upper()
            ):
                continue

            row_code = str(
                padded[3]
                or ""
            ).strip().upper()

            if (
                agent_code is not None
                and row_code
                != agent_code.upper()
            ):
                continue

            hour_text = str(
                padded[1]
                or ""
            ).strip()

            try:
                hour = int(
                    hour_text
                    .split(":")[0]
                )
            except (
                ValueError,
                IndexError,
            ):
                continue

            if not (
                0 <= hour <= 23
            ):
                continue

            spend_value = to_float(
                padded[5],
                default=float("nan"),
            )

            if not math.isfinite(
                spend_value
            ):
                continue

            if metric == "spend":
                value = spend_value

            elif metric == "percent":
                day_budget: float | None = None

                day_budget = (
                    TrackerSheet
                    .effective_budget_from_history(
                        budget_history,
                        date_value,
                        scope=scope,
                        agent_code=(
                            row_code
                            if scope.upper()
                            != "OVERALL"
                            else None
                        ),
                    )
                )

                if (
                    day_budget is not None
                    and day_budget > 0
                ):
                    value = (
                        spend_value
                        / day_budget
                    )

                else:
                    # Backward compatibility:
                    # use the raw percentage if an older row already has one.
                    pct_raw = padded[7]

                    if (
                        pct_raw is None
                        or str(
                            pct_raw
                        ).strip() == ""
                    ):
                        continue

                    value = to_float(
                        pct_raw,
                        default=float("nan"),
                    )

                    if not math.isfinite(
                        value
                    ):
                        continue

            else:
                raise ValueError(
                    f"Unsupported matrix metric: {metric}"
                )

            matrix[
                dt.day - 1
            ][hour] = value

        return matrix


    # --------------------------------------------------------
    # REFRESH VISIBLE MATRICES
    # --------------------------------------------------------

    def refresh_matrices(
        self,
        raw_rows: list[
            list[Any]
        ],
        now: datetime,
        agents: list[
            dict[str, str]
        ],
        budget_history: dict[
            str,
            dict[str, Any],
        ],
    ) -> None:

        raw_year = str(
            self.ws.acell(
                VIEW_YEAR_CELL
            ).value
            or ""
        ).strip()

        raw_month = str(
            self.ws.acell(
                VIEW_MONTH_CELL
            ).value
            or ""
        ).strip()

        extra_writes: list[
            dict[str, Any]
        ] = []

        try:
            year = int(
                float(
                    raw_year
                )
            )
        except ValueError:
            year = now.year

            extra_writes.append({
                "range": (
                    VIEW_YEAR_CELL
                ),
                "values": [[year]],
            })

        try:
            month = int(
                float(
                    raw_month
                )
            )

            if not (
                1 <= month <= 12
            ):
                raise ValueError

        except ValueError:
            month = now.month

            extra_writes.append({
                "range": (
                    VIEW_MONTH_CELL
                ),
                "values": [[month]],
            })

        overall_matrix = (
            self.build_matrix(
                raw_rows,
                year,
                month,
                scope="OVERALL",
                metric="percent",
                budget_history=(
                    budget_history
                ),
            )
        )

        overall_points = sum(
            1
            for matrix_row
            in overall_matrix
            for value
            in matrix_row
            if value != ""
        )

        print(
            "Overall percentage "
            "matrix points loaded: "
            f"{overall_points}"
        )

        spend_matrix = (
            self.build_matrix(
                raw_rows,
                year,
                month,
                scope="OVERALL",
                metric="spend",
                budget_history=(
                    budget_history
                ),
            )
        )

        spend_points = sum(
            1
            for matrix_row
            in spend_matrix
            for value
            in matrix_row
            if value != ""
        )

        print(
            "Overall spend "
            "matrix points loaded: "
            f"{spend_points}"
        )

        selected = str(
            self.ws.acell(
                AGENT_SELECTOR_CELL
            ).value
            or ""
        ).strip()

        parsed = (
            parse_agent_header(
                selected
            )
        )

        valid_codes = {
            agent["code"]
            for agent
            in agents
        }

        if (
            parsed
            and parsed[0]
            in valid_codes
        ):
            selected_code = (
                parsed[0]
            )

        else:
            selected_code = (
                agents[0]["code"]
            )

            extra_writes.append({
                "range": (
                    AGENT_SELECTOR_CELL
                ),
                "values": [[
                    agents[0][
                        "header"
                    ]
                ]],
            })

        agent_matrix = (
            self.build_matrix(
                raw_rows,
                year,
                month,
                scope="AGENT",
                metric="percent",
                budget_history=(
                    budget_history
                ),
                agent_code=(
                    selected_code
                ),
            )
        )

        agent_points = sum(
            1
            for matrix_row
            in agent_matrix
            for value
            in matrix_row
            if value != ""
        )

        print(
            "Agent percentage "
            f"matrix points loaded "
            f"({selected_code}): "
            f"{agent_points}"
        )

        self.batch_write([
            *extra_writes,
            {
                "range": (
                    OVERALL_MATRIX_RANGE
                ),
                "values": (
                    overall_matrix
                ),
            },
            {
                "range": (
                    SPEND_MATRIX_RANGE
                ),
                "values": (
                    spend_matrix
                ),
            },
            {
                "range": (
                    AGENT_MATRIX_RANGE
                ),
                "values": (
                    agent_matrix
                ),
            },
        ])


    # --------------------------------------------------------
    # KPI CARDS
    # --------------------------------------------------------

    def update_kpis(
        self,
        *,
        overall_spend: float,
        overall_budget: (
            float | None
        ),
        updated_at: str,
    ) -> None:

        if (
            overall_budget
            is not None
            and overall_budget > 0
        ):
            spend_ratio: (
                float | str
            ) = (
                overall_spend
                / overall_budget
            )

            remaining: (
                float | str
            ) = (
                overall_budget
                - overall_spend
            )
        else:
            spend_ratio = ""
            remaining = ""

        self.batch_write([
            {
                "range": (
                    LIVE_OVERALL_CELL
                ),
                "values": [[
                    spend_ratio
                ]],
            },
            {
                "range": (
                    SPEND_TODAY_CELL
                ),
                "values": [[
                    round(
                        overall_spend,
                        2,
                    )
                ]],
            },
            {
                "range": (
                    TODAY_BUDGET_CELL
                ),
                "values": [[
                    (
                        round(
                            overall_budget,
                            2,
                        )
                        if (
                            overall_budget
                            is not None
                        )
                        else "NO BUDGET"
                    )
                ]],
            },
            {
                "range": (
                    REMAINING_CELL
                ),
                "values": [[
                    (
                        round(
                            float(
                                remaining
                            ),
                            2,
                        )
                        if (
                            remaining
                            != ""
                        )
                        else ""
                    )
                ]],
            },
            {
                "range": (
                    LAST_UPDATE_CELL
                ),
                "values": [[
                    updated_at
                ]],
            },
        ])


    # --------------------------------------------------------
    # LIVE AGENT SNAPSHOT
    # --------------------------------------------------------

    def update_snapshot(
        self,
        agents: list[
            dict[str, str]
        ],
        spend_by_agent: dict[
            str,
            float,
        ],
        allocations: dict[
            str,
            float | None,
        ],
        expected_pace: float,
    ) -> None:

        output: list[
            list[Any]
        ] = []

        for agent in agents:

            code = agent["code"]

            spend = (
                spend_by_agent
                .get(
                    code,
                    0.0,
                )
            )

            budget = (
                allocations
                .get(code)
            )

            if (
                budget is not None
                and budget > 0
            ):

                ratio: (
                    float | str
                ) = (
                    spend
                    / budget
                )

                remaining: (
                    float | str
                ) = (
                    budget
                    - spend
                )

                delta: (
                    float | str
                ) = (
                    ratio
                    - expected_pace
                )

                if ratio > 1.05:
                    status = (
                        "OVER BUDGET"
                    )

                elif ratio >= 1.0:
                    status = (
                        "BUDGET REACHED"
                    )

                elif delta <= -0.10:
                    status = (
                        "BEHIND PACE"
                    )

                elif delta >= 0.10:
                    status = (
                        "AHEAD OF PACE"
                    )

                else:
                    status = (
                        "ON TRACK"
                    )

            else:

                ratio = ""
                remaining = ""
                delta = ""
                status = (
                    "NO BUDGET"
                )

            output.append([
                code,
                agent["name"],
                (
                    round(
                        budget,
                        2,
                    )
                    if (
                        budget
                        is not None
                    )
                    else ""
                ),
                round(
                    spend,
                    2,
                ),
                ratio,
                (
                    round(
                        float(
                            remaining
                        ),
                        2,
                    )
                    if remaining != ""
                    else ""
                ),
                expected_pace,
                delta,
                status,
            ])

        while len(
            output
        ) < 9:

            output.append(
                [""] * 9
            )

        self.batch_write([
            {
                "range": (
                    SNAPSHOT_RANGE
                ),
                "values": (
                    output[:9]
                ),
            },
        ])


# ============================================================
# AGENT MATCHING
# ============================================================

def extract_agent_code(
    account_name: str,
    agents: list[
        dict[str, str]
    ],
) -> str:

    text = str(
        account_name
        or ""
    ).upper()

    # Longer codes first.
    for agent in sorted(
        agents,
        key=lambda item: len(
            item["code"]
        ),
        reverse=True,
    ):

        code = (
            agent[
                "code"
            ].upper()
        )

        pattern = (
            rf"(?<![A-Z0-9])"
            rf"{re.escape(code)}"
            rf"(?![A-Z0-9])"
        )

        if re.search(
            pattern,
            text,
        ):
            return code

    return "UNKNOWN"


# ============================================================
# RAW ROW BUILDER
# ============================================================

def make_raw_row(
    *,
    today: str,
    hour_bucket: str,
    scope: str,
    code: str,
    name: str,
    spend: float,
    budget: float | None,
    updated_at: str,
) -> list[Any]:

    if (
        budget is not None
        and budget > 0
    ):

        spend_ratio: (
            float | str
        ) = (
            spend
            / budget
        )

        remaining: (
            float | str
        ) = (
            budget
            - spend
        )

    else:
        spend_ratio = ""
        remaining = ""

    return [
        today,
        hour_bucket,
        scope,
        code,
        name,
        round(
            spend,
            2,
        ),
        (
            round(
                budget,
                2,
            )
            if (
                budget
                is not None
            )
            else ""
        ),
        spend_ratio,
        (
            round(
                float(
                    remaining
                ),
                2,
            )
            if remaining != ""
            else ""
        ),
        updated_at,
    ]


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    timezone_name = (
        os.getenv(
            "REPORT_TIMEZONE",
            "Africa/Cairo",
        ).strip()
        or "Africa/Cairo"
    )

    now = datetime.now(
        ZoneInfo(
            timezone_name
        )
    )

    today = (
        now.date()
        .isoformat()
    )

    # Logical hourly bucket.
    #
    # Example:
    # actual run = 16:55
    # raw unique hour bucket = 16:00
    #
    # This means re-running at 16:57 updates the SAME 16:00 row
    # and never creates duplicate hourly snapshots.
    hour_bucket = (
        now.replace(
            minute=0,
            second=0,
            microsecond=0,
        )
        .strftime(
            "%H:%M"
        )
    )

    # Actual run time.
    # This is what is displayed in the automated hour header.
    updated_at = now.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    expected_pace = min(
        1.0,
        max(
            0.0,
            (
                (
                    now.hour
                    * 3600
                )
                + (
                    now.minute
                    * 60
                )
                + now.second
            )
            / 86400.0,
        ),
    )


    # --------------------------------------------------------
    # SHEET
    # --------------------------------------------------------

    sheet = TrackerSheet()

    agents = (
        sheet
        .read_agent_definitions()
    )

    # Keep internal tables aligned with visible agent headers.
    sheet.sync_internal_headers(
        agents
    )

    # One-time V6 design repair + visible table setup.
    sheet.ensure_visible_layout(
        agents
    )

    # AUTO HOURS — 12-hour format.
    sheet.sync_hour_headers(
        now
    )

    # Save the visible manual input row into historical budget table.
    saved_date = (
        sheet
        .save_input_to_history(
            agents,
            updated_at,
        )
    )

    # Read TODAY'S budget from historical table.
    budget = (
        sheet
        .read_budget_for_date(
            today,
            agents,
        )
    )

    # Load ALL dated budgets once for percentage backfill.
    budget_history = (
        sheet
        .read_budget_history(
            agents
        )
    )


    # --------------------------------------------------------
    # META
    # --------------------------------------------------------

    meta = MetaClient()

    accounts = (
        meta.discover_accounts()
    )

    (
        spend_rows,
        fetch_errors,
    ) = (
        meta
        .fetch_all_today_spend(
            accounts,
            today,
        )
    )

    overall_spend = sum(
        to_float(
            row.get(
                "spend"
            )
        )
        for row
        in spend_rows
    )


    # --------------------------------------------------------
    # AGENT SPEND
    # --------------------------------------------------------

    spend_by_agent = {
        agent["code"]: 0.0
        for agent
        in agents
    }

    unmapped: list[
        dict[str, Any]
    ] = []

    for row in spend_rows:

        spend = to_float(
            row.get(
                "spend"
            )
        )

        code = extract_agent_code(
            row[
                "account_name"
            ],
            agents,
        )

        if code == "UNKNOWN":

            if spend > 0:
                unmapped.append(
                    row
                )

            continue

        spend_by_agent[
            code
        ] = (
            spend_by_agent.get(
                code,
                0.0,
            )
            + spend
        )


    # --------------------------------------------------------
    # BUILD RAW HOURLY SNAPSHOT
    # --------------------------------------------------------

    raw_output: list[
        list[Any]
    ] = []

    # Overall first.
    raw_output.append(
        make_raw_row(
            today=today,
            hour_bucket=hour_bucket,
            scope="OVERALL",
            code="OVERALL",
            name="Overall",
            spend=overall_spend,
            budget=budget[
                "overall"
            ],
            updated_at=(
                updated_at
            ),
        )
    )

    # Then every agent.
    for agent in agents:

        raw_output.append(
            make_raw_row(
                today=today,
                hour_bucket=(
                    hour_bucket
                ),
                scope="AGENT",
                code=(
                    agent[
                        "code"
                    ]
                ),
                name=(
                    agent[
                        "name"
                    ]
                ),
                spend=(
                    spend_by_agent
                    .get(
                        agent[
                            "code"
                        ],
                        0.0,
                    )
                ),
                budget=(
                    budget[
                        "agents"
                    ].get(
                        agent[
                            "code"
                        ]
                    )
                ),
                updated_at=(
                    updated_at
                ),
            )
        )


    # --------------------------------------------------------
    # UPSERT RAW LOG
    # --------------------------------------------------------

    sheet.upsert_raw_rows(
        raw_output,
        today,
        hour_bucket,
    )

    # Read formatted raw rows.
    # Dates stay readable and percentage strings are parsed safely.
    all_raw = (
        sheet.raw_rows()
    )


    # --------------------------------------------------------
    # REFRESH MATRICES
    # --------------------------------------------------------

    sheet.refresh_matrices(
        all_raw,
        now,
        agents,
        budget_history,
    )


    # --------------------------------------------------------
    # REFRESH KPI CARDS
    # --------------------------------------------------------

    sheet.update_kpis(
        overall_spend=(
            overall_spend
        ),
        overall_budget=(
            budget[
                "overall"
            ]
        ),
        updated_at=(
            updated_at
        ),
    )


    # --------------------------------------------------------
    # REFRESH AGENT SNAPSHOT
    # --------------------------------------------------------

    sheet.update_snapshot(
        agents,
        spend_by_agent,
        budget["agents"],
        expected_pace,
    )


    # --------------------------------------------------------
    # LOGS
    # --------------------------------------------------------

    print(
        "=" * 90
    )

    print(
        "META HOURLY "
        "SPEND PACING"
    )

    print(
        f"Run: "
        f"{updated_at} "
        f"{timezone_name}"
    )

    print(
        "Actual header time: "
        f"{now.strftime('%H:%M')}"
    )

    print(
        "Raw hour bucket: "
        f"{hour_bucket}"
    )

    print(
        "Budget input saved date: "
        f"{saved_date or 'NONE'}"
    )

    print(
        "Today's budget found: "
        f"{budget['found']}"
    )

    print(
        "Budget source date: "
        f"{budget.get('source_date') or 'NONE'}"
        f" | inherited="
        f"{budget.get('inherited', False)}"
    )

    print(
        "Accounts discovered: "
        f"{len(accounts)}"
    )

    print(
        "Accounts fetched: "
        f"{len(spend_rows)}"
    )

    print(
        "Google Sheets writes: "
        "optimized with batch updates"
    )

    print(
        "=" * 90
    )

    overall_budget = (
        budget[
            "overall"
        ]
    )

    overall_pct = (
        (
            overall_spend
            / overall_budget
            * 100
        )
        if overall_budget
        else None
    )

    print(
        "OVERALL"
        f" | Spend="
        f"{overall_spend:,.2f}"
        f" | Budget="
        f"{overall_budget if overall_budget is not None else 'MISSING'}"
        f" | Used="
        f"{f'{overall_pct:.2f}%' if overall_pct is not None else 'N/A'}"
    )

    if fetch_errors:

        print(
            "\nMETA FETCH WARNINGS:"
        )

        for item in fetch_errors:
            print(
                f"- {item}"
            )

    if unmapped:

        print(
            "\nUNMAPPED ACCOUNTS "
            "WITH SPEND:"
        )

        for row in sorted(
            unmapped,
            key=lambda item: (
                -to_float(
                    item.get(
                        "spend"
                    )
                )
            ),
        ):

            print(
                f"- "
                f"{row['account_name']} "
                f"({row['id']})"
                f" | "
                f"{to_float(row.get('spend')):,.2f}"
            )

    print(
        "\nDone."
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )
