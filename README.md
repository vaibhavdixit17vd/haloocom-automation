# Haloocom Lead Extraction Automation (v2)

A hardened, fault-tolerant Python ETL job that incrementally extracts lead records from the Haloocom API, deduplicates overlapping pulls, exports clean Excel deliverables to a shared network folder, and notifies downstream teams via email.

---

## How It Works

Each execution follows a strict, crash-safe pipeline:

1. **Single-Instance Lock:** Creates an atomic `extract.lock` file containing the active PID to prevent overlapping cron/scheduled runs. Stale locks older than 3 hours are automatically cleared.
2. **Checkpoint Recovery:** Reads `state.json` to retrieve the last delivered timestamp (`watermark`) and the IDs of leads delivered within the previous overlap window (`seen`).
3. **Paginated Extraction:** Queries the Haloocom API starting from `watermark - OVERLAP` (oldest first) until an empty page is returned, with automatic retries and exponential backoff on `429` and `5xx` errors.
4. **Deduplication:** Removes duplicate lead IDs within the current pull and filters out any lead IDs already delivered in prior runs.
5. **Atomic File Delivery:** Writes records to a temporary `.part` file in the target shared folder before atomically renaming it to `Leads_YYYYMMDD_HHMMSS.xlsx` so downstream consumers never read a half-written file.
6. **State Commit & Notification:** Updates `state.json` atomically *immediately after* the Excel file is saved (preventing duplicate deliveries if SMTP fails), then sends summary emails to stakeholders—or an alert to the owner on failure.

---

## Prerequisites

- **Python:** 3.9+
- **Dependencies:** Install via `pip`:
  ```bash
  pip install -r requirements.txt



 ## Missing Details

 Configuration (Environment Variables)
All secrets, endpoints, paths, and SMTP settings are configured through environment variables.

API & Extraction Settings
Variable	Required	Default	Description
HALOOCOM_API_TOKEN	Yes	—	Bearer authentication token for the Haloocom API.
HALOOCOM_API_BASE	No	https://haloocom.internal/api	Base URL of the Haloocom API.
HALOOCOM_LEADS_ENDPOINT	No	/leads	Relative path for the leads endpoint.
HALOOCOM_API_TIMEZONE	No	UTC	Timezone assumed if API timestamps lack an explicit UTC offset.
OVERLAP_MINUTES	No	10	Lookback buffer subtracted from the watermark to catch delayed/late-arriving leads.
INITIAL_SINCE	No	1970-01-01T00:00:00+00:00	Starting ISO-8601 timestamp used only on the very first run when state.json does not exist.
Storage & Paths
Variable	Required	Default	Description
LEADS_SHARED_FOLDER	No	\\fileserver\shared\Leads	Destination directory where .xlsx extracts are delivered.
LEADS_DATA_DIR	No	Script directory	Directory for internal operational files (state.json, extract.lock, extract.log).
Email & Alerting (SMTP)
Variable	Required	Default	Description
SMTP_HOST	No	smtp.yourcompany.com	Mail server hostname.
SMTP_PORT	No	25	Mail server port.
SMTP_STARTTLS	No	false	Set to true to enable STARTTLS encryption.
SMTP_USER	No	""	SMTP authentication username (if required by server).
SMTP_PASSWORD	No	""	SMTP authentication password.
MAIL_FROM	No	lead-automation@yourcompany.com	Sender email address.
NOTIFY_TO	No	email-team@yourcompany.com,sv-team@yourcompany.com	Comma-separated list of recipients notified when a new extract is ready.
OWNER_EMAIL	No	you@yourcompany.com	Recipient notified if the extraction job encounters an unhandled error.
API Contract Verification
Before running in production, verify the constants in the API CONTRACT block at the top of the script against your Haloocom API documentation:

ID_FIELD ("id"): Unique identifier field on each lead object.

CREATED_FIELD ("created_at"): Creation timestamp field used for checkpointing and sorting.

RESULTS_KEY ("results"): JSON key containing the array of lead records.

PARAM_SINCE ("created_after"): Query parameter filtering leads created after a timestamp.

PARAM_SORT / SORT_VALUE ("sort", "created_at"): Must sort oldest-first (ascending) so pagination and watermarking remain consistent.

PARAM_PAGE / PARAM_PAGE_SIZE ("page", "page_size"): Pagination parameters (PAGE_SIZE = 500, MAX_PAGES = 200).

Usage & Scheduling
Manual Run
Set the required environment variables and execute the script:

Windows (PowerShell):

PowerShell
$env:HALOOCOM_API_TOKEN="your_secret_token"
$env:LEADS_SHARED_FOLDER="\\fileserver\shared\Leads"
python extract_leads.py
Linux / macOS:

Bash
export HALOOCOM_API_TOKEN="your_secret_token"
export LEADS_SHARED_FOLDER="/mnt/shared/Leads"
python3 extract_leads.py
Automated Scheduling
Windows Task Scheduler: Configure a recurring task (e.g., every 15 or 30 minutes) running python.exe with the script path as the argument.

Cron (Linux):

Code snippet
*/15 * * * * /usr/bin/python3 /opt/haloocom-extractor/extract_leads.py
Operational Files & Troubleshooting
All operational artifacts are stored in LEADS_DATA_DIR:

state.json: Tracks watermark (latest processed timestamp in UTC) and seen (map of recently delivered lead IDs). Note: If state.json becomes corrupted, the script intentionally halts rather than risking duplicate or missed data.

extract.log: Detailed execution logs, page counts, deduplication metrics, and stack traces.

extract.lock: Active process lock file. Automatically cleaned up on exit or after 3 hours if a previous process crashed hard.
