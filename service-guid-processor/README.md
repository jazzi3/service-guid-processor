# Service GUID Processor

This local web interface runs the same API request flow as `NG failed orders.py`.
It supports:

- Bearer token entry per processing session
- Editable organization, group, and source codes (defaulting to U002, D600, and D600)
- Eight-column table for Date, serviceguid, document_number, tenant_code, route_code, document_date, failure_reason, and Response
- Pasting tab-separated spreadsheet rows from the first Date cell; pasted values are read-only and Date plus serviceguid are required for every row
- One batch per table: clear the table to paste another batch; clearing cancels active processing and removes the current results
- Stops and highlights the token when the first GUID returns an upstream failure other than a network error, or when the API explicitly reports an authentication failure
- Marks the token field with a red Expired indicator when an authentication error is confirmed
- Live progress, cancellation, retry of failed IDs, and CSV result download

## Run on localhost (Windows)

### One-command startup

From Command Prompt, run:

```cmd
D:\Uzair\service-guid-processor\run.bat
```

Or, if Command Prompt is already open in this project folder, run:

```powershell
run.bat
```

The script creates `.venv` if needed, installs the dependencies, starts the
server, and opens the application in your browser. You can also press
`Ctrl+Shift+B` in VS Code and select **Run Service GUID Processor**.

The PowerShell equivalent is:

```powershell
.\run.ps1
```

If PowerShell blocks the script, allow local scripts once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Manual startup

Open PowerShell in this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Open http://127.0.0.1:8000 in a browser. When `NG failed orders.py` is present
next to this folder, the web app reads its exact API URL, including query
credentials, so both flows use the same endpoint. Otherwise configure
`API_URL`, `API_USERNAME`, and `API_PASSWORD` in `.env`. The bearer token must
also be the same token used by that script; save it in the page before starting
a job.

Paste one serviceGUID per line into the page. The web app sends the same POST URL, query parameters, JSON payload, and
`Authorization: Bearer <token>` header as the script. If the API returns
`"User is not allowed to perform this action"`, the API username/password or
bearer token does not match the working script.

## Free deployment on Render

This project includes `render.yaml` for a free Render web service. Render's
free service may sleep after inactivity, so the first request after a quiet
period can take a little longer. Jobs are kept in memory and are lost if the
service restarts; download results before ending a long session.

1. Put this project in a GitHub repository. Do not commit `.env` or real API
	credentials.
2. In Render, choose **New > Blueprint** and select the GitHub repository.
3. Confirm the service created from `render.yaml`.
4. In the service's environment settings, set `API_URL`. Set
	`API_USERNAME` and `API_PASSWORD` if the upstream API needs them.
5. Deploy and open the generated `https://...onrender.com` URL.

The deployed service uses the bearer token entered in the page. Render
provides HTTPS automatically, and the app's WebSocket progress connection
uses that secure connection automatically.
