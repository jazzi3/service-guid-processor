# Service GUID Processor

This local web interface runs the same API request flow as `NG failed orders.py`.
It supports:

- Bearer token entry per processing session
- Editable organization, group, and source codes (defaulting to U002, D600, and D600)
- Input table for Date, serviceguid, document_number, tenant_code, route_code, document_date, failure_reason, Status, and Response
- Pasting tab-separated spreadsheet rows from the first Date cell; pasted values are read-only and Date plus serviceguid are required for every row
- One batch per table: clear the table to paste another batch; clearing cancels active processing and removes the current results
- Stops and highlights the token when the first GUID returns an upstream failure other than a network error, or when the API explicitly reports an authentication failure
- Marks the token field with a red Expired indicator when an authentication error is confirmed
- Live progress, cancellation, retry of failed IDs, and CSV result download
- Automatic Google Sheets save after processing: new serviceGUIDs are appended, while existing GUID rows keep their metadata and have Status/Response refreshed
- Failure dashboard with inclusive Date From/Date To filters and a dynamic pivot table of trimmed failure reasons, daily totals, and grand totals
- Report view with date-filtered Google Sheets rows and an Excel workbook download
- Azure access-token generation using the local Azure CLI account

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

To use **Generate Azure Token**, install Azure CLI and run `az login` in CMD on
the computer whose browser you are using. Keep the local app running at
`http://127.0.0.1:8000`; when the page is opened from Render, the browser
requests a token from that local app, which runs
`az account get-access-token --resource-type oss-rdbms`. The local app accepts
this request only from the configured Render site or its own local page. The
generated token is separate from the service API bearer token.

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

### Configure automatic Google Sheets export

The exporter writes to the `Failed Orders` tab of the configured Google
spreadsheet using the Google Sheets API. Enable the Google Sheets API in a
Google Cloud project, create a service account, and grant its email address
**Editor** access to the spreadsheet. Public link access is not a substitute
for API authentication. Keep the service-account JSON private.

Set these variables in the Render service's environment settings (or in the
local `.env` file):

- `GOOGLE_SERVICE_ACCOUNT_JSON`: the complete service-account JSON credential
- `GOOGLE_SHEET_ID`: spreadsheet ID
- `GOOGLE_SHEET_NAME`: worksheet tab name (defaults to `Failed Orders`)

Never commit `.env`, `.env.local`, or put service-account credentials in
frontend code. The sheet must have a header row containing
`Date`, `serviceguid`, `document_number`,
`tenant_code`, `route_code`, `document_date`, `failure_reason`, `Reason`,
`Status`, and `Response`. If the sheet is completely blank, the exporter
creates this header row before saving the first batch. It adds missing
`Reason` and `Response` headers to an otherwise recognized sheet. Existing
`Reson` headers are renamed to `Reason` on the next save. It does not
add headers to a non-empty sheet with an unrecognized layout. Each completed
job appends new GUIDs and stores the API response.
When a processed GUID already exists, its Status and Response are refreshed
in place while the other existing cells are preserved. New records and
existing-row updates are sent together in a single Sheets batch update. The job
status reports export errors so failed saves are visible.
For new rows, `failure_reason` is kept as supplied and `Reason` is populated
from its JSON `message` field when present (or from the original text otherwise).

For local CSV testing, set `LOCAL_DATA_CSV_PATH` in the ignored `.env.local`
file. The local CSV is used for dashboard/report reads and processing-result
saves. Do not set this variable in Render; without it, the app uses the
configured Google Sheet.

Anyone-with-link editor access lets anyone with the link change the spreadsheet.
Restrict sharing to intended users and the service account if the data should
not be publicly editable.

The app uses the [Google Sheets API](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate)
to update the worksheet.

### Failure dashboard

Select **Dashboard** in the app to view failure counts by date. The Date From
and Date To filters default to today and update the pivot table inclusively.
The table displays a loading indicator while data is being retrieved.
Each distinct
`Reason` value (after trimming surrounding spaces) gets its own column, with
one row per distinct date, daily totals, and a grand-total row. Existing rows
with an empty `Reason` value use `failure_reason` as a fallback. The dashboard
reads from the configured data source (Google Sheets by default, or the local
CSV when configured) and refreshes after a processing batch is saved.

### Report

Select **Report** to see configured data-source rows for a date range. The
table fits the available width with long headers and values wrapping as
needed; the `Response` column is omitted from the on-screen table.
Both filters default to the current local date, and changing either filter
reloads the report with a loading indicator. **Download Excel** exports the
selected range as an `.xlsx` workbook.
