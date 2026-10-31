import asyncio
import csv
import io
import json
import os
import re
import tempfile
import threading
import uuid
from zipfile import BadZipFile
from copy import copy
from datetime import datetime
from typing import Any
from urllib.parse import quote
from xml.etree.ElementTree import ParseError

import requests
from google.auth.exceptions import GoogleAuthError
from google.auth.transport.requests import Request as GoogleAuthRequest
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException, WebSocket
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
API_URL = os.getenv("API_URL", "").strip()
API_USERNAME = os.getenv("API_USERNAME", "").strip()
API_PASSWORD = os.getenv("API_PASSWORD", "").strip()
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
GOOGLE_DRIVE_FILE_ID = os.getenv(
    "GOOGLE_DRIVE_FILE_ID",
    "18IowWJwC-_fUnJnAqPjKmHXRBQKbWQjU",
).strip()
REQUEST_TIMEOUT = (10, 30)   # (connect_timeout, read_timeout)
DELAY_SECONDS = 2
GOOGLE_DRIVE_API_URL = "https://www.googleapis.com/drive/v3"
GOOGLE_DRIVE_UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3"
GOOGLE_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"
WORKBOOK_COLUMNS = (
    "Date",
    "serviceguid",
    "document_number",
    "tenant_code",
    "route_code",
    "document_date",
    "failure_reason",
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


def _google_drive_error(response: requests.Response, action: str) -> WorkbookExportError:
    try:
        body = response.json()
    except ValueError:
        body = {}
    error = body.get("error") if isinstance(body, dict) else None
    message = error.get("message") if isinstance(error, dict) else None
    detail = f": {message}" if message else ""
    return WorkbookExportError(f"Google Drive {action} failed (HTTP {response.status_code}){detail}")


def _google_drive_access_token() -> str:
    missing = [
        name
        for name, value in (
            ("GOOGLE_SERVICE_ACCOUNT_JSON", GOOGLE_SERVICE_ACCOUNT_JSON),
            ("GOOGLE_DRIVE_FILE_ID", GOOGLE_DRIVE_FILE_ID),
        )
        if not value
    ]
    if missing:
        raise WorkbookExportError(
            "Google Drive export is not configured. Set: " + ", ".join(missing)
        )
    try:
        service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=[GOOGLE_DRIVE_SCOPE],
        )
        credentials.refresh(GoogleAuthRequest())
    except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise WorkbookExportError(
            "GOOGLE_SERVICE_ACCOUNT_JSON must contain valid Google service-account credentials."
        ) from exc
    except (GoogleAuthError, requests.exceptions.RequestException) as exc:
        raise WorkbookExportError(f"Google Drive authentication failed: {exc}") from exc
    if not credentials.token:
        raise WorkbookExportError("Google authentication returned no access token.")
    return credentials.token


def _google_drive_request(
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
        raise _google_drive_error(response, action)
    return response


def _normalise_header(value: Any) -> str:
    return re.sub(r"[\s_-]+", "", str(value or "").strip().casefold())


def _find_workbook_header(workbook: Any) -> tuple[Any, int, dict[str, int], bool]:
    required = {_normalise_header(name) for name in WORKBOOK_COLUMNS[:-1]}
    matches: list[tuple[Any, int, dict[str, int]]] = []
    for worksheet in workbook.worksheets:
        for row_number in range(1, worksheet.max_row + 1):
            headers: dict[str, int] = {}
            for column_number in range(1, worksheet.max_column + 1):
                header = _normalise_header(worksheet.cell(row_number, column_number).value)
                if header and header not in headers:
                    headers[header] = column_number
            if required.issubset(headers):
                matches.append((worksheet, row_number, headers))
    if not matches:
        has_existing_data = any(
            cell.value not in (None, "")
            for worksheet in workbook.worksheets
            for row in worksheet.iter_rows()
            for cell in row
        )
        if not has_existing_data:
            worksheet = workbook.active
            headers: dict[str, int] = {}
            for column_number, name in enumerate(WORKBOOK_COLUMNS, start=1):
                _set_text(worksheet.cell(1, column_number), name)
                headers[_normalise_header(name)] = column_number
            return worksheet, 1, headers, True
        expected = ", ".join(WORKBOOK_COLUMNS[:-1])
        raise WorkbookExportError(
            f"Could not find a worksheet row with the required headers: {expected}. "
            "Headers were not created because the workbook already contains data."
        )
    if len(matches) > 1:
        raise WorkbookExportError(
            "More than one worksheet header row matches the expected columns; no workbook changes were made."
        )
    worksheet, header_row, headers = matches[0]
    return worksheet, header_row, headers, False


def _set_text(cell: Any, value: Any) -> None:
    cell.value = "" if value is None else str(value)
    cell.data_type = "s"


def _table_for_header(worksheet: Any, header_row: int, columns: dict[str, int]) -> Any:
    required_columns = [
        columns[_normalise_header(name)] for name in WORKBOOK_COLUMNS[:-1]
    ]
    for table in worksheet.tables.values():
        min_column, min_row, max_column, _ = range_boundaries(table.ref)
        if (
            min_row == header_row
            and min_column <= min(required_columns)
            and max_column >= max(required_columns)
        ):
            return table
    return None


def append_results_to_google_drive(
    records: dict[str, dict[str, str]],
    results: list[dict[str, Any]],
) -> tuple[int, int]:
    token = _google_drive_access_token()
    file_id = quote(GOOGLE_DRIVE_FILE_ID, safe="")
    metadata_url = f"{GOOGLE_DRIVE_API_URL}/files/{file_id}"

    with _workbook_write_lock:
        try:
            metadata_response = _google_drive_request(
                "GET",
                metadata_url,
                token,
                "locate workbook",
                params={"fields": "id,name,mimeType,version"},
            )
            item = metadata_response.json()
        except ValueError as exc:
            raise WorkbookExportError("Google Drive returned invalid workbook metadata.") from exc
        try:
            item_id = item["id"]
            mime_type = item["mimeType"]
            file_version = item["version"]
        except (KeyError, TypeError) as exc:
            raise WorkbookExportError("Google Drive returned incomplete workbook metadata.") from exc
        if mime_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            raise WorkbookExportError(
                "The Google Drive file must be an uploaded .xlsx workbook, not a native Google Sheet or another file type."
            )
        content_url = f"{GOOGLE_DRIVE_API_URL}/files/{quote(item_id, safe='')}"
        content_response = _google_drive_request(
            "GET",
            content_url,
            token,
            "download workbook",
            params={"alt": "media"},
        )
        try:
            workbook = load_workbook(io.BytesIO(content_response.content))
        except (BadZipFile, InvalidFileException, OSError, ParseError, TypeError, ValueError) as exc:
            raise WorkbookExportError(
                "The Google Drive file could not be opened as a standard .xlsx workbook."
            ) from exc

        worksheet, header_row, headers, headers_created = _find_workbook_header(workbook)
        response_key = _normalise_header("Response")
        workbook_changed = headers_created or response_key not in headers
        response_added = response_key not in headers
        if response_added:
            response_column = max(headers.values()) + 1
            if response_column > 16384:
                raise WorkbookExportError("There is no available Excel column for the Response field.")
            headers[response_key] = response_column
            header_cell = worksheet.cell(header_row, response_column)
            _set_text(header_cell, "Response")
            status_column = headers.get(_normalise_header("Status"))
            if status_column:
                source_style = worksheet.cell(header_row, status_column)
                if source_style.has_style:
                    header_cell._style = copy(source_style._style)
                header_cell.number_format = source_style.number_format

        guid_column = headers[_normalise_header("serviceguid")]
        existing_rows: dict[str, int] = {}
        last_data_row = header_row
        for row_number in range(header_row + 1, worksheet.max_row + 1):
            row_has_data = any(
                worksheet.cell(row_number, column_number).value not in (None, "")
                for column_number in range(1, worksheet.max_column + 1)
            )
            if row_has_data:
                last_data_row = row_number
            guid_value = worksheet.cell(row_number, guid_column).value
            if guid_value not in (None, ""):
                existing_rows.setdefault(str(guid_value).strip().casefold(), row_number)

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
            output_values = {
                _normalise_header("Status"): (
                    "Successfully" if result.get("Status") == "Success" else "Failed"
                ),
                response_key: result.get("Response", ""),
            }
            for normalized_header, value in output_values.items():
                _set_text(worksheet.cell(target_row, headers[normalized_header]), value)
            workbook_changed = True

        for metadata, result in new_rows:
            target_row = last_data_row + 1
            if target_row > 1_048_576:
                raise WorkbookExportError("The worksheet has no remaining rows for new records.")
            output_values = {
                _normalise_header("Date"): metadata.get("date", ""),
                _normalise_header("serviceguid"): result.get("serviceGUID", ""),
                _normalise_header("document_number"): metadata.get("document_number", ""),
                _normalise_header("tenant_code"): metadata.get("tenant_code", ""),
                _normalise_header("route_code"): metadata.get("route_code", ""),
                _normalise_header("document_date"): metadata.get("document_date", ""),
                _normalise_header("failure_reason"): metadata.get("failure_reason", ""),
                _normalise_header("Status"): (
                    "Successfully" if result.get("Status") == "Success" else "Failed"
                ),
                response_key: result.get("Response", ""),
            }
            for normalized_header, value in output_values.items():
                column_number = headers[normalized_header]
                cell = worksheet.cell(target_row, column_number)
                if last_data_row > header_row:
                    source = worksheet.cell(last_data_row, column_number)
                    if source.has_style:
                        cell._style = copy(source._style)
                    cell.number_format = source.number_format
                _set_text(cell, value)
            last_data_row = target_row
            workbook_changed = True

        table = _table_for_header(worksheet, header_row, headers)
        if table and workbook_changed:
            min_column, min_row, max_column, _ = range_boundaries(table.ref)
            new_max_column = max(max_column, headers[response_key])
            if new_max_column > max_column:
                if headers[response_key] != max_column + 1:
                    raise WorkbookExportError(
                        "The Response column must immediately follow the existing Excel table columns."
                    )
                from openpyxl.worksheet.table import TableColumn

                table.tableColumns.append(
                    TableColumn(id=len(table.tableColumns) + 1, name="Response")
                )
            table.ref = (
                f"{get_column_letter(min_column)}{min_row}:"
                f"{get_column_letter(new_max_column)}{last_data_row}"
            )
            if table.autoFilter:
                table.autoFilter.ref = table.ref

        if workbook_changed:
            output = io.BytesIO()
            try:
                workbook.save(output)
            except (OSError, TypeError, ValueError) as exc:
                raise WorkbookExportError("The updated Excel workbook could not be prepared for upload.") from exc

            upload_headers = {
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
            etag = metadata_response.headers.get("ETag")
            if etag:
                upload_headers["If-Match"] = etag
            else:
                try:
                    latest = _google_drive_request(
                        "GET",
                        metadata_url,
                        token,
                        "verify workbook has not changed",
                        params={"fields": "version"},
                    ).json()
                except ValueError as exc:
                    raise WorkbookExportError(
                        "Google Drive returned invalid workbook version metadata."
                    ) from exc
                if str(latest.get("version")) != str(file_version):
                    raise WorkbookExportError(
                        "The workbook changed while records were being prepared; no upload was made. Retry the save."
                    )
            _google_drive_request(
                "PATCH",
                f"{GOOGLE_DRIVE_UPLOAD_URL}/files/{quote(item_id, safe='')}",
                token,
                "save workbook",
                params={"uploadType": "media"},
                headers=upload_headers,
                data=output.getvalue(),
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
        job["status_message"] = "Processing complete. Saving records to Google Drive..."
        try:
            added, updated = append_results_to_google_drive(job.get("records", {}), job["results"])
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
