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


# Overall matrix
# A13 = Day
# B13:Y13 = hourly headers
# B14:Y44 = 31 days x 24 hours
OVERALL_DAY_HEADER_CELL = "A13"
OVERALL_HOUR_HEADER_RANGE = "B13:Y13"
OVERALL_MATRIX_RANGE = "B14:Y44"


# Agent matrix
# B48 = selected agent
# A50 = Day
# B50:Y50 = hourly headers
# B51:Y81 = 31 days x 24 hours
AGENT_SELECTOR_CELL = "B48"
AGENT_DAY_HEADER_CELL = "A50"
AGENT_HOUR_HEADER_RANGE = "B50:Y50"
AGENT_MATRIX_RANGE = "B51:Y81"


# Live agents snapshot
SNAPSHOT_RANGE = "A86:I94"


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

        book = gc.open_by_key(
            required_env(
                "GOOGLE_SHEET_ID"
            )
        )

        try:
            self.ws = book.worksheet(
                SHEET_TAB
            )

        except gspread.WorksheetNotFound:

            raise RuntimeError(
                f"Tab '{SHEET_TAB}' "
                "was not found. "
                "Upload the provided "
                "template first."
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

        self.ws.update(
            HISTORY_HEADER_RANGE,
            [history_headers],
        )

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

        self.ws.update(
            RAW_HEADER_RANGE,
            [raw_headers],
        )


    # --------------------------------------------------------
    # AUTOMATED HOUR HEADERS
    # --------------------------------------------------------

    def sync_hour_headers(
        self,
        now: datetime,
    ) -> None:
        """
        Hour labels are maintained automatically.

        Initial/default:
            00:00, 01:00, 02:00 ... 23:00

        On every run, only the CURRENT hour gets the real update time.

        Example:
            Run at 16:55
            -> hour 16 label becomes 16:55

            Next scheduled run at 17:05
            -> hour 17 label becomes 17:05

        Previous hour labels are preserved.

        This keeps the same 24-column matrix, while showing when
        each hourly snapshot was actually refreshed.
        """

        default_headers = [
            f"{hour:02d}:00"
            for hour in range(24)
        ]

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

            current_value = str(
                existing[hour]
                or ""
            ).strip()

            headers.append(
                current_value
                or default_headers[
                    hour
                ]
            )

        # Update current hour only
        headers[
            now.hour
        ] = now.strftime(
            "%H:%M"
        )

        self.ws.update(
            OVERALL_DAY_HEADER_CELL,
            [["Day"]],
        )

        self.ws.update(
            OVERALL_HOUR_HEADER_RANGE,
            [headers],
            value_input_option=(
                "USER_ENTERED"
            ),
        )

        self.ws.update(
            AGENT_DAY_HEADER_CELL,
            [["Day"]],
        )

        self.ws.update(
            AGENT_HOUR_HEADER_RANGE,
            [headers],
            value_input_option=(
                "USER_ENTERED"
            ),
        )


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

            self.ws.update(
                INPUT_STATUS_CELL,
                [["NO DATE"]],
            )

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

            self.ws.update(
                INPUT_STATUS_CELL,
                [[
                    "NO BUDGETS ENTERED"
                ]],
            )

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

        if target_row is None:

            target_row = (
                HISTORY_DATA_START
                + len(
                    history_values
                )
            )

        self.ws.update(
            (
                f"AA{target_row}"
                f":AM{target_row}"
            ),
            [output],
            value_input_option=(
                "USER_ENTERED"
            ),
        )

        self.ws.update(
            INPUT_STATUS_CELL,
            [[
                f"SAVED "
                f"{budget_date}"
            ]],
        )

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

        rows = self.ws.get(
            f"AA"
            f"{HISTORY_DATA_START}"
            f":AM"
        )

        matched: list[
            Any
        ] | None = None

        for row in rows:

            padded = (
                list(row)
                + [""] * (
                    13
                    - len(row)
                )
            )

            if (
                parse_sheet_date(
                    padded[0]
                )
                == target_date
            ):
                matched = padded

        allocations: dict[
            str,
            float | None,
        ] = {
            agent["code"]: None
            for agent
            in agents
        }

        overall: float | None = (
            None
        )

        if matched is not None:

            overall_raw = str(
                matched[1]
                or ""
            ).strip()

            if overall_raw:

                candidate = to_float(
                    overall_raw,
                    default=float(
                        "nan"
                    ),
                )

                if (
                    math.isfinite(
                        candidate
                    )
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
                    matched[idx]
                    or ""
                ).strip()

                if not raw_value:
                    continue

                candidate = to_float(
                    raw_value,
                    default=float(
                        "nan"
                    ),
                )

                if (
                    math.isfinite(
                        candidate
                    )
                    and candidate >= 0
                ):
                    allocations[
                        agent["code"]
                    ] = candidate

        valid_agent_budgets = [
            value
            for value
            in allocations.values()
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
                matched
                is not None
            ),
        }


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
        Unique hourly key:
            Date + Hour Bucket + Agent Code

        So multiple runs in the same hour UPDATE the existing row
        instead of creating duplicates.
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
            row,
        ) in enumerate(
            existing,
            start=RAW_DATA_START,
        ):

            padded = (
                list(row)
                + [""] * (
                    10
                    - len(row)
                )
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

                self.ws.update(
                    (
                        f"AO{target}"
                        f":AX{target}"
                    ),
                    [row],
                    value_input_option=(
                        "USER_ENTERED"
                    ),
                )

            else:
                append_rows.append(
                    row
                )

        if not append_rows:
            return

        next_row = (
            RAW_DATA_START
            + len(existing)
        )

        for row in append_rows:

            self.ws.update(
                (
                    f"AO{next_row}"
                    f":AX{next_row}"
                ),
                [row],
                value_input_option=(
                    "USER_ENTERED"
                ),
            )

            next_row += 1


    def raw_rows(
        self,
    ) -> list[list[Any]]:
        """
        IMPORTANT:
        Request UNFORMATTED values from Google Sheets.

        Therefore a 66.41% cell normally comes back as 0.6641,
        not the text "66.41%".

        to_float() still also supports the formatted string as backup.
        """

        return self.ws.get(
            f"AO"
            f"{RAW_DATA_START}"
            f":AX",
            value_render_option=(
                "UNFORMATTED_VALUE"
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
        agent_code: (
            str | None
        ) = None,
    ) -> list[list[Any]]:

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

            if (
                agent_code
                is not None
            ):

                row_code = str(
                    padded[3]
                    or ""
                ).strip().upper()

                if (
                    row_code
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

            pct_raw = padded[7]

            if (
                pct_raw is None
                or str(
                    pct_raw
                ).strip() == ""
            ):
                continue

            pct_value = to_float(
                pct_raw,
                default=float(
                    "nan"
                ),
            )

            if not math.isfinite(
                pct_value
            ):
                continue

            matrix[
                dt.day - 1
            ][hour] = pct_value

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

        try:
            year = int(
                float(
                    raw_year
                )
            )

        except ValueError:

            year = now.year

            self.ws.update(
                VIEW_YEAR_CELL,
                [[year]],
            )

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

            self.ws.update(
                VIEW_MONTH_CELL,
                [[month]],
            )

        overall_matrix = (
            self.build_matrix(
                raw_rows,
                year,
                month,
                scope="OVERALL",
            )
        )

        self.ws.update(
            OVERALL_MATRIX_RANGE,
            overall_matrix,
            value_input_option=(
                "USER_ENTERED"
            ),
        )

        selected = str(
            self.ws.acell(
                AGENT_SELECTOR_CELL
            ).value
            or ""
        ).strip()

        parsed = parse_agent_header(
            selected
        )

        if parsed:
            selected_code = (
                parsed[0]
            )

        else:
            selected_code = (
                agents[0]["code"]
            )

            self.ws.update(
                AGENT_SELECTOR_CELL,
                [[
                    agents[0][
                        "header"
                    ]
                ]],
            )

        agent_matrix = (
            self.build_matrix(
                raw_rows,
                year,
                month,
                scope="AGENT",
                agent_code=(
                    selected_code
                ),
            )
        )

        self.ws.update(
            AGENT_MATRIX_RANGE,
            agent_matrix,
            value_input_option=(
                "USER_ENTERED"
            ),
        )


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

        values = {
            LIVE_OVERALL_CELL: (
                spend_ratio
            ),
            SPEND_TODAY_CELL: round(
                overall_spend,
                2,
            ),
            TODAY_BUDGET_CELL: (
                round(
                    overall_budget,
                    2,
                )
                if (
                    overall_budget
                    is not None
                )
                else "NO BUDGET"
            ),
            REMAINING_CELL: (
                round(
                    float(
                        remaining
                    ),
                    2,
                )
                if remaining != ""
                else ""
            ),
            LAST_UPDATE_CELL: (
                updated_at
            ),
        }

        for (
            cell,
            value,
        ) in values.items():

            self.ws.update(
                cell,
                [[value]],
                value_input_option=(
                    "USER_ENTERED"
                ),
            )

        try:

            self.ws.format(
                "G5:J5",
                {
                    "numberFormat": {
                        "type": (
                            "PERCENT"
                        ),
                        "pattern": (
                            "0.00%"
                        ),
                    }
                },
            )

            self.ws.format(
                "K5:V5",
                {
                    "numberFormat": {
                        "type": (
                            "NUMBER"
                        ),
                        "pattern": (
                            "#,##0.00"
                        ),
                    }
                },
            )

        except Exception:
            # Formatting is cosmetic.
            # Never fail the run for it.
            pass


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

        self.ws.update(
            SNAPSHOT_RANGE,
            output[:9],
            value_input_option=(
                "USER_ENTERED"
            ),
        )


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

    # AUTO HOURS
    # No manual editing for 00:00 / 01:00 / etc.
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

    # Read raw rows as UNFORMATTED values.
    # This fixes Google % parsing.
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
        "Accounts discovered: "
        f"{len(accounts)}"
    )

    print(
        "Accounts fetched: "
        f"{len(spend_rows)}"
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
