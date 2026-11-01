import asyncio
import csv
import io
import json
import os
import re
import tempfile
import threading
import uuid
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any
from urllib.parse import quote

from openpyxl import Workbook
import requests
from google.auth.exceptions import GoogleAuthError
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.service_account import Credentials
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException, WebSocket
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
load_dotenv(os.path.join(BASE_DIR, ".env.local"), override=True)
API_URL = os.getenv("API_URL", "").strip()
API_USERNAME = os.getenv("API_USERNAME", "").strip()
API_PASSWORD = os.getenv("API_PASSWORD", "").strip()
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
GOOGLE_SHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID",
    "1H6rbt2woyWUFBRVlqGBSVI1mIJ4EaBZawSIksgmXJKE",
).strip()
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_NAME", "Failed Orders").strip()
LOCAL_DATA_CSV_PATH = os.getenv("LOCAL_DATA_CSV_PATH", "").strip()
REQUEST_TIMEOUT = (10, 30)   # (connect_timeout, read_timeout)
DELAY_SECONDS = 2
GOOGLE_SHEETS_API_URL = "https://sheets.googleapis.com/v4"
GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
WORKBOOK_COLUMNS = (
    "Date",
    "serviceguid",
    "document_number",
    "tenant_code",
    "route_code",
    "document_date",
    "failure_reason",
    "Reason",
    "Status",
    "Response",
)
_workbook_write_lock = threading.Lock()

# Retry up to 3 times on connection-level failures (RemoteDisconnected, reset, etc.).
# backoff_factor=1 means waits: 0s, 2s, 4s between retries.
_RETRY = Retry(
    total=3,
    backoff_factor=1,
    allowed_methods={"POST"},
    status_forcelist={502, 503, 504},
    raise_on_status=False,
)


def _make_session() -> requests.Session:
    """Return a Session that retries on transient connection errors."""
    session = requests.Session()
    adapter = HTTPAdapter(max_retries=_RETRY)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

app = FastAPI(title="Service GUID Processor")
app.mount("/templates", StaticFiles(directory=os.path.join(BASE_DIR, "templates")), name="templates")
jobs: dict[str, dict[str, Any]] = {}


class WorkbookExportError(Exception):
    pass


def _google_sheets_error(response: requests.Response, action: str) -> WorkbookExportError:
    try:
        body = response.json()
    except ValueError:
        body = {}
    error = body.get("error") if isinstance(body, dict) else None
    message = error.get("message") if isinstance(error, dict) else None
    detail = f": {message}" if message else ""
    return WorkbookExportError(f"Google Sheets {action} failed (HTTP {response.status_code}){detail}")


def _google_sheets_access_token() -> str:
    missing = [
        name
        for name, value in (
            ("GOOGLE_SERVICE_ACCOUNT_JSON", GOOGLE_SERVICE_ACCOUNT_JSON),
            ("GOOGLE_SHEET_ID", GOOGLE_SHEET_ID),
            ("GOOGLE_SHEET_NAME", GOOGLE_SHEET_NAME),
        )
        if not value
    ]
    if missing:
        raise WorkbookExportError(
            "Google Sheets export is not configured. Set: " + ", ".join(missing)
        )
    try:
        service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=[GOOGLE_SHEETS_SCOPE],
        )
        credentials.refresh(GoogleAuthRequest())
    except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise WorkbookExportError(
            "GOOGLE_SERVICE_ACCOUNT_JSON must contain valid Google service-account credentials."
        ) from exc
    except (GoogleAuthError, requests.exceptions.RequestException) as exc:
        raise WorkbookExportError(f"Google Sheets authentication failed: {exc}") from exc
    if not credentials.token:
        raise WorkbookExportError("Google authentication returned no access token.")
    return credentials.token


def _google_sheets_request(
    method: str,
    url: str,
    token: str,
    action: str,
    **kwargs: Any,
) -> requests.Response:
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    headers.setdefault("Accept", "application/json")
    try:
        response = requests.request(method, url, headers=headers, timeout=60, **kwargs)
    except requests.exceptions.RequestException as exc:
        raise WorkbookExportError(f"Could not {action}: {exc}") from exc
    if not response.ok:
        raise _google_sheets_error(response, action)
    return response


def _normalise_header(value: Any) -> str:
    return re.sub(r"[\s_-]+", "", str(value or "").strip().casefold())


def _find_sheet_header(
    values: list[list[Any]],
    required_columns: tuple[str, ...] | None = None,
) -> tuple[int, dict[str, int], bool]:
    if required_columns is None:
        optional_columns = {
            _normalise_header("Reason"),
            _normalise_header("Reson"),
            _normalise_header("Response"),
        }
        required_columns = tuple(
            name for name in WORKBOOK_COLUMNS
            if _normalise_header(name) not in optional_columns
        )
    required = {_normalise_header(name) for name in required_columns}
    matches: list[tuple[int, dict[str, int]]] = []
    for row_index, row in enumerate(values):
        headers: dict[str, int] = {}
        for column_index, value in enumerate(row):
            header = _normalise_header(value)
            if header and header not in headers:
                headers[header] = column_index
        if required.issubset(headers):
            matches.append((row_index, headers))
    if not matches:
        has_existing_data = any(
            value not in (None, "")
            for row in values
            for value in row
        )
        if not has_existing_data:
            return 0, {
                _normalise_header(name): column_index
                for column_index, name in enumerate(WORKBOOK_COLUMNS)
            }, True
        expected = ", ".join(required_columns)
        raise WorkbookExportError(
            f"Could not find a row with the required headers: {expected}. "
            "Headers were not created because the sheet already contains data."
        )
    if len(matches) > 1:
        raise WorkbookExportError(
            "More than one row matches the expected headers; no sheet changes were made."
        )
    header_row, headers = matches[0]
    return header_row, headers, False


def _sheet_cell(value: Any) -> dict[str, dict[str, str]]:
    return {"userEnteredValue": {"stringValue": "" if value is None else str(value)}}


def _parse_sheet_date(value: Any, row_number: int) -> date:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return date(1899, 12, 30) + timedelta(days=int(value))
        except (OverflowError, ValueError) as exc:
            raise WorkbookExportError(
                f"Invalid date serial in Google Sheet row {row_number}."
            ) from exc
    if isinstance(value, str):
        text = value.strip()
        try:
            return date.fromisoformat(text[:10])
        except ValueError as exc:
            raise WorkbookExportError(
                f"Invalid Date in Google Sheet row {row_number}: {text!r}. "
                "Use an ISO date (YYYY-MM-DD) or a date-formatted cell."
            ) from exc
    raise WorkbookExportError(
        f"Missing or invalid Date in Google Sheet row {row_number}."
    )


def _build_failure_pivot(
    values: list[list[Any]],
    date_from: date | None,
    date_to: date | None,
) -> dict[str, Any]:
    header_row, headers, headers_created = _find_sheet_header(values, ("Date",))
    if headers_created:
        return {
            "reasons": [],
            "rows": [],
            "grand_totals": {},
            "overall_total": 0,
            "record_count": 0,
            "min_date": "",
            "max_date": "",
        }

    date_column = headers[_normalise_header("Date")]
    reason_column = headers.get(
        _normalise_header("Reason"),
        headers.get(
            _normalise_header("Reson"),
            headers.get(_normalise_header("failure_reason"), -1),
        ),
    )
    if reason_column < 0:
        raise WorkbookExportError(
            'The worksheet must contain a "Reason" or "failure_reason" column.'
        )
    failure_reason_column = headers.get(_normalise_header("failure_reason"), -1)
    counts_by_date: dict[date, dict[str, int]] = {}
    for row_index in range(header_row + 1, len(values)):
        row = values[row_index]
        if not any(value not in (None, "") for value in row):
            continue
        raw_date = row[date_column] if date_column < len(row) else ""
        if raw_date in (None, ""):
            raise WorkbookExportError(
                f"Missing Date in Google Sheet row {row_index + 1}."
            )
        current_date = _parse_sheet_date(raw_date, row_index + 1)
        if date_from and current_date < date_from:
            continue
        if date_to and current_date > date_to:
            continue
        raw_reason = row[reason_column] if reason_column < len(row) else ""
        if raw_reason in (None, "") and failure_reason_column >= 0:
            raw_reason = (
                row[failure_reason_column]
                if failure_reason_column < len(row)
                else ""
            )
        reason = _extract_failure_message(raw_reason).strip()
        if failure_reason_column >= 0:
            raw_failure_reason = (
                row[failure_reason_column]
                if failure_reason_column < len(row)
                else ""
            )
            extracted_failure = _extract_failure_message(raw_failure_reason).strip()
            if extracted_failure and extracted_failure in reason and extracted_failure != reason:
                reason = extracted_failure
        reason = reason or "(Blank)"
        date_counts = counts_by_date.setdefault(current_date, {})
        date_counts[reason] = date_counts.get(reason, 0) + 1

    reasons = sorted(
        {reason for daily_counts in counts_by_date.values() for reason in daily_counts},
        key=lambda reason: (reason.casefold(), reason),
    )
    grand_totals = {
        reason: sum(daily_counts.get(reason, 0) for daily_counts in counts_by_date.values())
        for reason in reasons
    }
    rows = [
        {
            "date": current_date.isoformat(),
            "counts": daily_counts,
            "total": sum(daily_counts.values()),
        }
        for current_date, daily_counts in sorted(counts_by_date.items())
    ]
    return {
        "reasons": reasons,
        "rows": rows,
        "grand_totals": grand_totals,
        "overall_total": sum(grand_totals.values()),
        "record_count": sum(row["total"] for row in rows),
        "min_date": rows[0]["date"] if rows else "",
        "max_date": rows[-1]["date"] if rows else "",
    }


def _read_google_sheet_values(action: str) -> list[list[Any]]:
    if LOCAL_DATA_CSV_PATH:
        if not os.path.isfile(LOCAL_DATA_CSV_PATH):
            raise WorkbookExportError(
                f"Local CSV data file does not exist: {LOCAL_DATA_CSV_PATH}"
            )
        try:
            with open(LOCAL_DATA_CSV_PATH, encoding="utf-8-sig", newline="") as csv_file:
                return list(csv.reader(csv_file))
        except (OSError, UnicodeError, csv.Error) as exc:
            raise WorkbookExportError(
                f"Could not {action} from the configured local CSV: {exc}"
            ) from exc

    spreadsheet_id = quote(GOOGLE_SHEET_ID, safe="")
    spreadsheet_url = f"{GOOGLE_SHEETS_API_URL}/spreadsheets/{spreadsheet_id}"
    escaped_sheet_name = GOOGLE_SHEET_NAME.replace("'", "''")
    values_range = quote(f"'{escaped_sheet_name}'!A:AZ", safe="!:$'")
    try:
        token = _google_sheets_access_token()
        spreadsheet = _google_sheets_request(
            "GET",
            spreadsheet_url,
            token,
            "read spreadsheet metadata",
            params={"fields": "sheets(properties(title))"},
        ).json()
        if not any(
            item.get("properties", {}).get("title") == GOOGLE_SHEET_NAME
            for item in spreadsheet.get("sheets", [])
        ):
            raise WorkbookExportError(
                f'The spreadsheet does not contain a tab named "{GOOGLE_SHEET_NAME}".'
            )
        sheet_data = _google_sheets_request(
            "GET",
            f"{spreadsheet_url}/values/{values_range}",
            token,
            action,
            params={
                "valueRenderOption": "UNFORMATTED_VALUE",
                "dateTimeRenderOption": "SERIAL_NUMBER",
            },
        ).json()
    except ValueError as exc:
        raise WorkbookExportError("Google Sheets returned invalid worksheet data.") from exc
    return sheet_data.get("values", [])


def _build_sheet_report(
    values: list[list[Any]],
    date_from: date,
    date_to: date,
) -> tuple[list[str], list[list[Any]]]:
    header_row, headers, headers_created = _find_sheet_header(
        values,
        ("Date", "serviceguid"),
    )
    if headers_created:
        return list(WORKBOOK_COLUMNS), []

    date_column = headers[_normalise_header("Date")]
    reason_column = headers.get(
        _normalise_header("Reason"),
        headers.get(_normalise_header("Reson"), -1),
    )
    failure_reason_column = headers.get(_normalise_header("failure_reason"), -1)
    source_headers = values[header_row]
    column_count = max(
        len(source_headers),
        max((len(row) for row in values[header_row + 1 :]), default=0),
    )
    report_headers = [
        str(source_headers[index]).strip()
        if index < len(source_headers) and str(source_headers[index]).strip()
        else f"Column {index + 1}"
        for index in range(column_count)
    ]
    for column_index, header in enumerate(report_headers):
        if _normalise_header(header) == _normalise_header("Reson"):
            report_headers[column_index] = "Reason"
    report_rows: list[list[Any]] = []
    for row_index in range(header_row + 1, len(values)):
        row = values[row_index]
        if not any(value not in (None, "") for value in row):
            continue
        raw_date = row[date_column] if date_column < len(row) else ""
        if raw_date in (None, ""):
            raise WorkbookExportError(
                f"Missing Date in Google Sheet row {row_index + 1}."
            )
        current_date = _parse_sheet_date(raw_date, row_index + 1)
        if current_date < date_from or current_date > date_to:
            continue
        report_row = [
            row[index] if index < len(row) else ""
            for index in range(column_count)
        ]
        report_row[date_column] = current_date.isoformat()
        if reason_column >= 0 and failure_reason_column >= 0:
            reason = _extract_failure_message(report_row[reason_column]).strip()
            failure_reason = _extract_failure_message(
                report_row[failure_reason_column]
            ).strip()
            if failure_reason and failure_reason in reason and failure_reason != reason:
                report_row[reason_column] = failure_reason
        report_rows.append(report_row)
    return report_headers, report_rows


def _load_filtered_report(
    date_from: date | None,
    date_to: date | None,
) -> tuple[date, date, list[str], list[list[Any]]]:
    current_date = date.today()
    start_date = date_from or current_date
    end_date = date_to or current_date
    if start_date > end_date:
        raise HTTPException(
            status_code=400,
            detail="Date From must be on or before Date To.",
        )
    try:
        values = _read_google_sheet_values("read report data")
        headers, rows = _build_sheet_report(values, start_date, end_date)
    except WorkbookExportError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    return start_date, end_date, headers, rows


def _append_results_to_csv(
    csv_path: str,
    records: dict[str, dict[str, str]],
    results: list[dict[str, Any]],
) -> tuple[int, int]:
    if not os.path.isfile(csv_path):
        raise WorkbookExportError(f"Local CSV data file does not exist: {csv_path}")

    with _workbook_write_lock:
        try:
            with open(csv_path, encoding="utf-8-sig", newline="") as csv_file:
                rows = list(csv.reader(csv_file))
        except (OSError, UnicodeError, csv.Error) as exc:
            raise WorkbookExportError(f"Could not read local CSV data: {exc}") from exc

        header_row, headers, headers_created = _find_sheet_header(rows)
        if headers_created:
            rows = [list(WORKBOOK_COLUMNS)]
            header_row = 0
            headers = {
                _normalise_header(name): index
                for index, name in enumerate(WORKBOOK_COLUMNS)
            }

        reason_key = _normalise_header("Reason")
        legacy_reason_key = _normalise_header("Reson")
        if reason_key not in headers:
            if legacy_reason_key in headers:
                reason_column = headers[legacy_reason_key]
                rows[header_row][reason_column] = "Reason"
                headers[reason_key] = reason_column
            else:
                reason_column = max(headers.values(), default=-1) + 1
                headers[reason_key] = reason_column
                if len(rows[header_row]) <= reason_column:
                    rows[header_row] += [""] * (reason_column + 1 - len(rows[header_row]))
                rows[header_row][reason_column] = "Reason"

        response_key = _normalise_header("Response")
        if response_key not in headers:
            response_column = max(headers.values(), default=-1) + 1
            headers[response_key] = response_column
            if len(rows[header_row]) <= response_column:
                rows[header_row] += [""] * (response_column + 1 - len(rows[header_row]))
            rows[header_row][response_column] = "Response"

        column_count = max(headers.values(), default=-1) + 1
        rows[header_row] += [""] * (column_count - len(rows[header_row]))
        for row in rows[header_row + 1 :]:
            row += [""] * (column_count - len(row))

        guid_column = headers[_normalise_header("serviceguid")]
        existing_rows: dict[str, int] = {}
        for row_index in range(header_row + 1, len(rows)):
            guid = rows[row_index][guid_column].strip()
            if guid:
                existing_rows.setdefault(guid.casefold(), row_index)

        new_rows: list[tuple[dict[str, str], dict[str, Any]]] = []
        seen_guids: set[str] = set()
        updated = 0
        for result in results:
            guid = str(result.get("serviceGUID", "")).strip()
            normalized_guid = guid.casefold()
            if not normalized_guid or normalized_guid in seen_guids:
                continue
            seen_guids.add(normalized_guid)
            existing_row = existing_rows.get(normalized_guid)
            if existing_row is not None:
                rows[existing_row][headers[_normalise_header("Status")]] = (
                    "Successfully" if result.get("Status") == "Success" else "Failed"
                )
                rows[existing_row][headers[response_key]] = str(
                    result.get("Response", "") or ""
                )
                updated += 1
            else:
                new_rows.append((records.get(normalized_guid, {}), result))

        for metadata, result in new_rows:
            row = [""] * column_count
            output_values = {
                _normalise_header("Date"): metadata.get("date", ""),
                _normalise_header("serviceguid"): result.get("serviceGUID", ""),
                _normalise_header("document_number"): metadata.get("document_number", ""),
                _normalise_header("tenant_code"): metadata.get("tenant_code", ""),
                _normalise_header("route_code"): metadata.get("route_code", ""),
                _normalise_header("document_date"): metadata.get("document_date", ""),
                _normalise_header("failure_reason"): metadata.get("failure_reason", ""),
                reason_key: _extract_failure_message(metadata.get("failure_reason", "")),
                _normalise_header("Status"): (
                    "Successfully" if result.get("Status") == "Success" else "Failed"
                ),
                response_key: result.get("Response", ""),
            }
            for normalized_header, value in output_values.items():
                row[headers[normalized_header]] = "" if value is None else str(value)
            rows.append(row)

        temporary_path = ""
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                newline="",
                dir=os.path.dirname(os.path.abspath(csv_path)),
                delete=False,
            ) as csv_file:
                temporary_path = csv_file.name
                csv.writer(csv_file).writerows(rows)
                csv_file.flush()
                os.fsync(csv_file.fileno())
            os.replace(temporary_path, csv_path)
        except OSError as exc:
            if temporary_path and os.path.exists(temporary_path):
                os.remove(temporary_path)
            raise WorkbookExportError(f"Could not save local CSV data: {exc}") from exc

        return len(new_rows), updated


def append_results_to_google_sheet(
    records: dict[str, dict[str, str]],
    results: list[dict[str, Any]],
) -> tuple[int, int]:
    if LOCAL_DATA_CSV_PATH:
        return _append_results_to_csv(LOCAL_DATA_CSV_PATH, records, results)

    token = _google_sheets_access_token()
    spreadsheet_id = quote(GOOGLE_SHEET_ID, safe="")
    spreadsheet_url = f"{GOOGLE_SHEETS_API_URL}/spreadsheets/{spreadsheet_id}"
    escaped_sheet_name = GOOGLE_SHEET_NAME.replace("'", "''")
    values_range = quote(f"'{escaped_sheet_name}'!A:AZ", safe="!:$'")

    with _workbook_write_lock:
        try:
            spreadsheet = _google_sheets_request(
                "GET",
                spreadsheet_url,
                token,
                "read spreadsheet metadata",
                params={"fields": "sheets(properties(sheetId,title))"},
            ).json()
        except ValueError as exc:
            raise WorkbookExportError("Google Sheets returned invalid spreadsheet metadata.") from exc
        sheets = spreadsheet.get("sheets", [])
        sheet = next(
            (
                item.get("properties", {})
                for item in sheets
                if item.get("properties", {}).get("title") == GOOGLE_SHEET_NAME
            ),
            None,
        )
        if sheet is None:
            raise WorkbookExportError(
                f'The spreadsheet does not contain a tab named "{GOOGLE_SHEET_NAME}".'
            )
        try:
            sheet_id = sheet["sheetId"]
        except KeyError as exc:
            raise WorkbookExportError("Google Sheets returned incomplete tab metadata.") from exc

        try:
            sheet_data = _google_sheets_request(
                "GET",
                f"{spreadsheet_url}/values/{values_range}",
                token,
                "read worksheet data",
            ).json()
        except ValueError as exc:
            raise WorkbookExportError("Google Sheets returned invalid worksheet data.") from exc
        values = sheet_data.get("values", [])
        header_row, headers, headers_created = _find_sheet_header(values)
        response_key = _normalise_header("Response")
        reason_key = _normalise_header("Reason")
        legacy_reason_key = _normalise_header("Reson")
        requests_to_apply: list[dict[str, Any]] = []
        if headers_created:
            requests_to_apply.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "startColumnIndex": 0,
                        },
                        "rows": [
                            {"values": [_sheet_cell(name) for name in WORKBOOK_COLUMNS]}
                        ],
                        "fields": "userEnteredValue",
                    }
                }
            )
        if reason_key not in headers:
            if legacy_reason_key in headers:
                reason_column = headers[legacy_reason_key]
                headers[reason_key] = reason_column
                requests_to_apply.append(
                    {
                        "updateCells": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": header_row,
                                "startColumnIndex": reason_column,
                            },
                            "rows": [{"values": [_sheet_cell("Reason")]}],
                            "fields": "userEnteredValue",
                        }
                    }
                )
            else:
                reason_column = max(headers.values(), default=-1) + 1
                headers[reason_key] = reason_column
                if not headers_created:
                    requests_to_apply.append(
                        {
                            "updateCells": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": header_row,
                                    "startColumnIndex": reason_column,
                                },
                                "rows": [{"values": [_sheet_cell("Reason")]}],
                                "fields": "userEnteredValue",
                            }
                        }
                    )

        for optional_header in ("Response",):
            normalized_header = _normalise_header(optional_header)
            if normalized_header not in headers:
                column_index = max(headers.values(), default=-1) + 1
                headers[normalized_header] = column_index
                if not headers_created:
                    requests_to_apply.append(
                        {
                            "updateCells": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": header_row,
                                    "startColumnIndex": column_index,
                                },
                                "rows": [{"values": [_sheet_cell(optional_header)]}],
                                "fields": "userEnteredValue",
                            }
                        }
                    )

        guid_column = headers[_normalise_header("serviceguid")]
        existing_rows: dict[str, int] = {}
        for row_index in range(header_row + 1, len(values)):
            row = values[row_index]
            guid_value = row[guid_column] if guid_column < len(row) else ""
            if guid_value not in (None, ""):
                existing_rows.setdefault(str(guid_value).strip().casefold(), row_index)

        new_rows: list[tuple[dict[str, str], dict[str, Any]]] = []
        existing_updates: list[tuple[int, dict[str, Any]]] = []
        seen_guids: set[str] = set()
        updated = 0
        for result in results:
            guid = str(result.get("serviceGUID", "")).strip()
            normalized_guid = guid.casefold()
            if not normalized_guid or normalized_guid in seen_guids:
                continue
            seen_guids.add(normalized_guid)
            existing_row = existing_rows.get(normalized_guid)
            if existing_row is not None:
                existing_updates.append((existing_row, result))
                updated += 1
                continue
            new_rows.append((records.get(normalized_guid, {}), result))

        for target_row, result in existing_updates:
            status_value = "Successfully" if result.get("Status") == "Success" else "Failed"
            for header, value in (
                ("Status", status_value),
                ("Response", result.get("Response", "")),
            ):
                requests_to_apply.append(
                    {
                        "updateCells": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": target_row,
                                "startColumnIndex": headers[_normalise_header(header)],
                            },
                            "rows": [{"values": [_sheet_cell(value)]}],
                            "fields": "userEnteredValue",
                        }
                    }
                )

        row_width = max(headers.values(), default=-1) + 1
        appended_rows = []
        for metadata, result in new_rows:
            output_values = {
                _normalise_header("Date"): metadata.get("date", ""),
                _normalise_header("serviceguid"): result.get("serviceGUID", ""),
                _normalise_header("document_number"): metadata.get("document_number", ""),
                _normalise_header("tenant_code"): metadata.get("tenant_code", ""),
                _normalise_header("route_code"): metadata.get("route_code", ""),
                _normalise_header("document_date"): metadata.get("document_date", ""),
                _normalise_header("failure_reason"): metadata.get("failure_reason", ""),
                reason_key: _extract_failure_message(
                    metadata.get("failure_reason", "")
                ),
                _normalise_header("Status"): (
                    "Successfully" if result.get("Status") == "Success" else "Failed"
                ),
                response_key: result.get("Response", ""),
            }
            row = [""] * row_width
            for normalized_header, value in output_values.items():
                row[headers[normalized_header]] = "" if value is None else str(value)
            appended_rows.append({"values": [_sheet_cell(value) for value in row]})

        if appended_rows:
            requests_to_apply.append(
                {
                    "appendCells": {
                        "sheetId": sheet_id,
                        "rows": appended_rows,
                        "fields": "userEnteredValue",
                    }
                }
            )
        if requests_to_apply:
            _google_sheets_request(
                "POST",
                f"{spreadsheet_url}:batchUpdate",
                token,
                "save worksheet changes",
                json={"requests": requests_to_apply},
            )
        return len(new_rows), updated


def configured_api_url() -> str:
    script_path = os.path.join(os.path.dirname(BASE_DIR), "NG failed orders.py")
    try:
        with open(script_path, encoding="utf-8") as script:
            match = re.search(r'^\s*URL\s*=\s*["\']([^"\']+)["\']', script.read(), re.MULTILINE)
    except OSError:
        match = None
    return match.group(1).strip() if match else API_URL


def empty_result(guid: str, status: str, response: str, http_code: Any = "") -> dict[str, Any]:
    return {
        "serviceGUID": guid,
        "Status": status,
        "HTTP Code": http_code,
        "Response": response,
        "DisplayResponse": clean_response(response),
    }


def is_authentication_error(http_code: Any, response: str) -> bool:
    if http_code in {401, 403}:
        return True
    message = re.sub(r"[\s_-]+", " ", response.lower())
    return any(
        phrase in message
        for phrase in (
            "token expired",
            "token has expired",
            "token has been expired",
            "token is expired",
            "expired token",
            "invalid token",
            "token is invalid",
            "unauthorized",
            "not authorized",
            "authentication failed",
            "authentication token",
        )
    )


def is_network_error(result: dict[str, Any]) -> bool:
    return not result["HTTP Code"] and result["Response"].startswith("REQUEST ERROR:")


def is_failed_without_response(result: dict[str, Any]) -> bool:
    if is_network_error(result):
        return False
    response = result["Response"].strip().lower()
    return not response or response == "no response"


def clean_response(response: str) -> str:
    try:
        data = json.loads(response)
    except (TypeError, ValueError):
        return response[:500]
    if isinstance(data, dict):
        for key in ("description", "message", "error", "status"):
            value = data.get(key)
            if value not in (None, ""):
                return str(value)
    return str(data)[:500]


def _extract_failure_message(value: Any) -> str:
    text = "" if value is None else str(value)
    try:
        data = json.loads(text)
    except ValueError:
        decoder = json.JSONDecoder()
        data = None
        for offset, character in enumerate(text):
            if character != "{":
                continue
            try:
                data, _ = decoder.raw_decode(text[offset:])
                break
            except ValueError:
                continue
    if isinstance(data, dict):
        for key in ("message", "error"):
            message = data.get(key)
            if message not in (None, ""):
                return str(message)
    if isinstance(data, str):
        return data
    return text


def parse_guids(text: str) -> list[str]:
    seen: set[str] = set()
    values: list[str] = []
    for line in text.splitlines():
        guid = line.strip()
        key = guid.casefold()
        if guid and key not in seen:
            seen.add(key)
            values.append(guid)
    return values


def parse_record_metadata(text: str, guids: list[str]) -> dict[str, dict[str, str]]:
    if not text:
        return {}
    try:
        rows = json.loads(text)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Pasted row metadata must be valid JSON.") from exc
    if not isinstance(rows, list) or not all(isinstance(row, dict) for row in rows):
        raise HTTPException(status_code=400, detail="Pasted row metadata must be a list of rows.")
    valid_guids = {guid.casefold() for guid in guids}
    fields = ("date", "serviceguid", "document_number", "tenant_code", "route_code", "document_date", "failure_reason")
    metadata: dict[str, dict[str, str]] = {}
    for row in rows:
        guid = str(row.get("serviceguid", "")).strip()
        key = guid.casefold()
        if not key or key not in valid_guids or key in metadata:
            continue
        metadata[key] = {field: str(row.get(field, "") or "") for field in fields if field != "serviceguid"}
    return metadata


def build_payload(
    service_guid: str,
    organization_code: str = "U002",
    group_code: str = "D600",
    source_code: str = "D600",
) -> dict[str, Any]:
    return {
        "organizationCode": organization_code,
        "groupCode": group_code,
        "sourceCode": source_code,
        "job": None,
        "sourceQueryParams": {"serviceGUID": service_guid},
        "identification": "node",
        "event": "Node_Execution",
    }


def run_job(job_id: str) -> None:
    job = jobs[job_id]
    job["state"] = "processing"

    # Resolve the API URL once — avoids repeated disk I/O inside the loop.
    api_url = configured_api_url()
    if not api_url:
        job["state"] = "completed"
        job["status_message"] = "Error: API_URL is not configured in .env."
        return

    headers = {
        "Authorization": f"Bearer {job['token']}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        # Disable keep-alive so each request gets a fresh connection.
        # This prevents RemoteDisconnected errors when the server closes
        # an idle pooled connection between requests.
        "Connection": "close",
    }
    request_params: dict[str, str] = {}
    if "?" not in api_url:
        if API_USERNAME:
            request_params["user"] = API_USERNAME
        if API_PASSWORD:
            request_params["password"] = API_PASSWORD

    with _make_session() as session:
        for position, guid in enumerate(job["guids"]):
            if job["cancel_event"].is_set():
                break
            job["current"] = guid
            job["status_message"] = "Sending API request..."
            try:
                response = session.post(
                    api_url,
                    params=request_params,
                    headers=headers,
                    json=build_payload(
                        guid,
                        job["organization_code"],
                        job["group_code"],
                        job["source_code"],
                    ),
                    timeout=REQUEST_TIMEOUT,
                )
                status = "Success" if response.ok else "Failed"
                result = empty_result(guid, status, response.text, response.status_code)
            except requests.exceptions.RequestException as exc:
                result = empty_result(guid, "Failed", f"REQUEST ERROR: {exc}", "")
            except Exception as exc:
                result = empty_result(guid, "Failed", f"UNEXPECTED ERROR: {exc}", "")

            auth_expired = is_authentication_error(result["HTTP Code"], result["Response"])
            if position == 0 and result["Status"] == "Failed" and not is_network_error(result):
                auth_expired = True
            if is_failed_without_response(result):
                auth_expired = True
            if auth_expired:
                result["Status"] = "Failed"
            with job["lock"]:
                job["results"].append(result)
                job["completed"] += 1
                if result["Status"] == "Success":
                    job["successful"] += 1
                else:
                    job["failed"] += 1
            if auth_expired:
                job["auth_expired"] = True
                job["status_message"] = "Dcode Token Has Been Expired. Processing stopped."
                break
            job["status_message"] = "Request completed."
            if position < len(job["guids"]) - 1 and not job["cancel_event"].wait(DELAY_SECONDS):
                continue
            if job["cancel_event"].is_set():
                break

    job["current"] = ""
    with job["lock"]:
        save_to_workbook = not job["cancel_event"].is_set() and bool(job["results"])
        if save_to_workbook:
            job["workbook_status"] = "saving"
    if save_to_workbook:
        job["status_message"] = "Processing complete. Saving records to Google Sheets..."
        try:
            added, updated = append_results_to_google_sheet(job.get("records", {}), job["results"])
            job["workbook_status"] = "saved"
            job["workbook_added"] = added
            job["workbook_updated"] = updated
        except WorkbookExportError as exc:
            job["workbook_status"] = "error"
            job["workbook_error"] = str(exc)
    job["state"] = "cancelled" if job["cancel_event"].is_set() else "completed"
    if job["state"] == "cancelled":
        job["status_message"] = "Processing cancelled."
    elif not job["auth_expired"]:
        job["status_message"] = "Processing complete."

def create_job(
    token: str,
    guids: list[str],
    records: dict[str, dict[str, str]] | None = None,
    organization_code: str = "U002",
    group_code: str = "D600",
    source_code: str = "D600",
) -> str:
    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "id": job_id,
        "token": token,
        "guids": guids,
        "records": records or {},
        "organization_code": organization_code,
        "group_code": group_code,
        "source_code": source_code,
        "results": [],
        "state": "starting",
        "status_message": "Starting...",
        "auth_expired": False,
        "workbook_status": "pending",
        "workbook_added": 0,
        "workbook_updated": 0,
        "workbook_error": "",
        "current": "",
        "completed": 0,
        "successful": 0,
        "failed": 0,
        "cancel_event": threading.Event(),
        "lock": threading.Lock(),
        "created": datetime.now().isoformat(),
    }
    threading.Thread(target=run_job, args=(job_id,), daemon=True).start()
    return job_id


@app.get("/", response_class=HTMLResponse)
async def index() -> FileResponse:
    return FileResponse(os.path.join(BASE_DIR, "templates", "index.html"))


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/dashboard")
def failure_dashboard(
    date_from: date | None = None,
    date_to: date | None = None,
) -> dict[str, Any]:
    if date_from and date_to and date_from > date_to:
        raise HTTPException(status_code=400, detail="Date From must be on or before Date To.")

    try:
        values = _read_google_sheet_values("read dashboard data")
    except WorkbookExportError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc

    try:
        return _build_failure_pivot(values, date_from, date_to)
    except WorkbookExportError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc


@app.get("/api/report")
def sheet_report(
    date_from: date | None = None,
    date_to: date | None = None,
) -> dict[str, Any]:
    start_date, end_date, headers, rows = _load_filtered_report(date_from, date_to)
    return {
        "headers": headers,
        "rows": rows,
        "record_count": len(rows),
        "date_from": start_date.isoformat(),
        "date_to": end_date.isoformat(),
    }


@app.get("/api/report/export")
def export_sheet_report(
    date_from: date | None = None,
    date_to: date | None = None,
) -> Response:
    start_date, end_date, headers, rows = _load_filtered_report(date_from, date_to)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Report"
    worksheet.append(headers)
    _, sheet_headers, _ = _find_sheet_header(
        [headers],
        ("Date", "serviceguid"),
    )
    date_column = sheet_headers[_normalise_header("Date")]
    for row in rows:
        output_row = list(row)
        output_row[date_column] = date.fromisoformat(output_row[date_column])
        worksheet.append(output_row)
        worksheet.cell(worksheet.max_row, date_column + 1).number_format = "yyyy-mm-dd"
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    buffer = BytesIO()
    workbook.save(buffer)
    filename = f"service-guid-report-{start_date.isoformat()}-to-{end_date.isoformat()}.xlsx"
    return Response(
        content=buffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/jobs")
async def start_job(
    bearer_token: str = Form(...),
    guids: str = Form(""),
    records: str = Form(""),
    organization_code: str = Form("U002"),
    group_code: str = Form("D600"),
    source_code: str = Form("D600"),
) -> dict[str, Any]:
    token = bearer_token.strip()
    if not token:
        raise HTTPException(status_code=400, detail="Bearer Token is required.")
    values = parse_guids(guids)
    if not values:
        raise HTTPException(status_code=400, detail="Enter at least one serviceGUID.")
    record_metadata = parse_record_metadata(records, values)
    codes = [organization_code.strip(), group_code.strip(), source_code.strip()]
    if not all(codes):
        raise HTTPException(status_code=400, detail="Organization, group, and source codes are required.")
    job_id = create_job(
        token,
        values,
        records=record_metadata,
        organization_code=codes[0],
        group_code=codes[1],
        source_code=codes[2],
    )
    return {"job_id": job_id, "total": len(values)}


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str) -> dict[str, str]:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    with job["lock"]:
        if job["workbook_status"] == "saving":
            raise HTTPException(status_code=409, detail="Processing is finished; the Google Drive save cannot be cancelled.")
        job["cancel_event"].set()
    return {"message": "Cancellation requested."}


@app.post("/api/jobs/{job_id}/retry")
async def retry_job(job_id: str) -> dict[str, Any]:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    failed_guids = [row["serviceGUID"] for row in job["results"] if row["Status"] == "Failed"]
    if not failed_guids:
        raise HTTPException(status_code=400, detail="There are no failed GUIDs to retry.")
    retry_id = create_job(
        job["token"],
        failed_guids,
        records={
            guid.casefold(): job.get("records", {}).get(guid.casefold(), {})
            for guid in failed_guids
        },
        organization_code=job["organization_code"],
        group_code=job["group_code"],
        source_code=job["source_code"],
    )
    return {"job_id": retry_id, "total": len(failed_guids)}


@app.get("/api/jobs/{job_id}")
async def job_status(job_id: str) -> dict[str, Any]:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    # Snapshot results inside the lock to avoid a race with the writer thread.
    with job["lock"]:
        results_snapshot = list(job["results"])
    return {
        "id": job["id"],
        "state": job["state"],
        "status_message": job["status_message"],
        "current": job["current"],
        "total": len(job["guids"]),
        "completed": job["completed"],
        "successful": job["successful"],
        "failed": job["failed"],
        "remaining": len(job["guids"]) - job["completed"],
        "auth_expired": job["auth_expired"],
        "workbook_status": job["workbook_status"],
        "workbook_added": job["workbook_added"],
        "workbook_updated": job["workbook_updated"],
        "workbook_error": job["workbook_error"],
        "results": results_snapshot,
    }


@app.websocket("/ws/jobs/{job_id}")
async def job_socket(websocket: WebSocket, job_id: str) -> None:
    if job_id not in jobs:
        await websocket.close(code=1008)
        return
    await websocket.accept()
    try:
        while True:
            await websocket.send_json(await job_status(job_id))
            if jobs[job_id]["state"] in {"completed", "cancelled"}:
                break
            await asyncio.sleep(0.4)
    finally:
        await websocket.close()


@app.get("/api/jobs/{job_id}/download")
async def download_results(job_id: str) -> FileResponse:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    # Snapshot results to avoid holding the lock while doing file I/O.
    with job["lock"]:
        results_snapshot = list(job["results"])
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["serviceGUID", "Status", "HTTP Code", "Response"])
    writer.writeheader()
    # extrasaction='ignore' silently drops the DisplayResponse key that is
    # present in results but not in the CSV fieldnames.
    writer = csv.DictWriter(
        output,
        fieldnames=["serviceGUID", "Status", "HTTP Code", "Response"],
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(results_snapshot)
    # Write to a named temp file so we don't accumulate files in results/.
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".csv", prefix=f"service-guid-results-{job_id}-",
        delete=False, newline="", encoding="utf-8",
    )
    tmp.write(output.getvalue())
    tmp.close()
    return FileResponse(
        tmp.name,
        filename=f"service-guid-results-{job_id}.csv",
        media_type="text/csv",
        background=None,
    )
