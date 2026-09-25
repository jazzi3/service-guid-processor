import asyncio
import csv
import io
import os
import threading
import uuid
from datetime import datetime
import re
from typing import Any

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Form, HTTPException, WebSocket
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
API_URL = os.getenv("API_URL", "").strip()
API_USERNAME = os.getenv("API_USERNAME", "").strip()
API_PASSWORD = os.getenv("API_PASSWORD", "").strip()
REQUEST_TIMEOUT = 30
DELAY_SECONDS = 2

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
        data = __import__("json").loads(response)
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


def build_payload(service_guid: str) -> dict[str, Any]:
    return {
        "organizationCode": "U002",
        "groupCode": "D600",
        "sourceCode": "D600",
        "job": None,
        "sourceQueryParams": {"serviceGUID": service_guid},
        "identification": "node",
        "event": "Node_Execution",
    }


def run_job(job_id: str) -> None:
    job = jobs[job_id]
    job["state"] = "processing"
    headers = {
        "Authorization": f"Bearer {job['token']}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    for position, guid in enumerate(job["guids"]):
        if job["cancel_event"].is_set():
            break
        job["current"] = guid
        job["status_message"] = "Sending API request..."
        try:
            api_url = configured_api_url()
            if not api_url:
                raise RuntimeError("API_URL is not configured in .env.")
            request_params = {}
            if "?" not in api_url:
                if API_USERNAME:
                    request_params["user"] = API_USERNAME
                if API_PASSWORD:
                    request_params["password"] = API_PASSWORD
            response = requests.post(
                api_url,
                params=request_params,
                headers=headers,
                json=build_payload(guid),
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


def create_job(token: str, guids: list[str]) -> str:
    job_id = uuid.uuid4().hex
    jobs[job_id] = {
        "id": job_id,
        "token": token,
        "guids": guids,
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
) -> dict[str, Any]:
    token = bearer_token.strip()
    if not token:
        raise HTTPException(status_code=400, detail="Bearer Token is required.")
    values = parse_guids(guids)
    if not values:
        raise HTTPException(status_code=400, detail="Enter at least one serviceGUID.")
    job_id = create_job(token, values)
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
    retry_id = create_job(job["token"], failed_guids)
    return {"job_id": retry_id, "total": len(failed_guids)}


@app.get("/api/jobs/{job_id}")
async def job_status(job_id: str) -> dict[str, Any]:
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found.")
    return {
        "id": job["id"], "state": job["state"], "status_message": job["status_message"],
        "current": job["current"], "total": len(job["guids"]), "completed": job["completed"],
        "successful": job["successful"], "failed": job["failed"],
        "remaining": len(job["guids"]) - job["completed"], "results": job["results"],
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
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["serviceGUID", "Status", "HTTP Code", "Response"])
    writer.writeheader()
    writer.writerows(job["results"])
    results_dir = os.path.join(BASE_DIR, "results")
    path = os.path.join(results_dir, f"service-guid-results-{job_id}.csv")
    os.makedirs(results_dir, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as destination:
        destination.write(output.getvalue())
    return FileResponse(path, filename=os.path.basename(path), media_type="text/csv")
