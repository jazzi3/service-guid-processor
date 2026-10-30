import asyncio
import csv
import io
import json
import os
import re
import tempfile
import threading
import uuid
from datetime import datetime
from typing import Any

import requests
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
REQUEST_TIMEOUT = (10, 30)   # (connect_timeout, read_timeout)
DELAY_SECONDS = 2

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
        if guid and guid not in seen:
            seen.add(guid)
            values.append(guid)
    return values


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

            with job["lock"]:
                job["results"].append(result)
                job["completed"] += 1
                if result["Status"] == "Success":
                    job["successful"] += 1
                else:
                    job["failed"] += 1
            job["status_message"] = "Request completed."
            if position < len(job["guids"]) - 1 and not job["cancel_event"].wait(DELAY_SECONDS):
                continue
            if job["cancel_event"].is_set():
                break

    job["current"] = ""
    job["state"] = "cancelled" if job["cancel_event"].is_set() else "completed"
    job["status_message"] = "Processing cancelled." if job["state"] == "cancelled" else "Processing complete."

def create_job(
    token: str,
    guids: list[str],
    organization_code: str = "U002",
    group_code: str = "D600",
    source_code: str = "D600",
) -> str:
    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "id": job_id,
        "token": token,
        "guids": guids,
        "organization_code": organization_code,
        "group_code": group_code,
        "source_code": source_code,
        "results": [],
        "state": "starting",
        "status_message": "Starting...",
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
    codes = [organization_code.strip(), group_code.strip(), source_code.strip()]
    if not all(codes):
        raise HTTPException(status_code=400, detail="Organization, group, and source codes are required.")
    job_id = create_job(token, values, *codes)
    return {"job_id": job_id, "total": len(values)}


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str) -> dict[str, str]:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
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
        job["organization_code"],
        job["group_code"],
        job["source_code"],
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
