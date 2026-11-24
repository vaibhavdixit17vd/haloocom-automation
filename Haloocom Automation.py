"""
Haloocom lead extraction automation (v2, hardened).

Each run:
 1. Takes a single-instance lock (no overlapping runs).
 2. Reads the checkpoint (watermark = newest lead created_at already delivered
    + IDs of leads delivered inside the overlap window).
 3. Pulls leads created after (watermark - overlap), oldest first, page by page.
 4. Removes duplicates by lead ID and drops leads already delivered earlier.
 5. Writes the Excel file to the shared folder (temp name, then atomic rename).
 6. Saves the new checkpoint, THEN notifies the teams by email.

Setup:  pip install requests pandas openpyxl
Secrets and paths come from environment variables.

!! Confirm every item in the "API CONTRACT" block against the Haloocom docs. !!
"""
import pandas as pd
import json
import logging
import os
import smtplib
import time
from contextlib import contextmanager
from datetime import timedelta
from email.message import EmailMessage
from pathlib import Path


import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def _env_list(name: str, default: str = "") -> list:
    return [x.strip() for x in os.environ.get(name, default).split(",") if x.strip()]


# ------------------------------ CONFIG ------------------------------
API_BASE = os.environ.get("HALOOCOM_API_BASE", "https://haloocom.internal/api").rstrip("/")
API_TOKEN = os.environ["HALOOCOM_API_TOKEN"]
LEADS_ENDPOINT = os.environ.get("HALOOCOM_LEADS_ENDPOINT", "/leads")

# ---- API CONTRACT (placeholders: CONFIRM against Haloocom API docs) ----
ID_FIELD = "id"                    # unique lead identifier in each record
CREATED_FIELD = "created_at"       # lead creation timestamp in each record
RESULTS_KEY = "results"            # key holding the list of leads in the response
PARAM_SINCE = "created_after"      # filter: only leads created after this time
PARAM_SORT, SORT_VALUE = "sort", "created_at"   # must be OLDEST FIRST (ascending)
PARAM_PAGE, PARAM_PAGE_SIZE = "page", "page_size"
PAGE_SIZE = 500
MAX_PAGES = 200                    # safety stop against runaway loops
API_TIMEZONE = os.environ.get("HALOOCOM_API_TIMEZONE", "UTC")  # for timestamps without offset
# -------------------------------------------------------------------------

OVERLAP = timedelta(minutes=int(os.environ.get("OVERLAP_MINUTES", "10")))
INITIAL_SINCE = os.environ.get("INITIAL_SINCE", "1970-01-01T00:00:00+00:00")  # first run only

SHARED_FOLDER = Path(os.environ.get("LEADS_SHARED_FOLDER", r"\\fileserver\shared\Leads"))
DATA_DIR = Path(os.environ.get("LEADS_DATA_DIR", Path(__file__).resolve().parent))
STATE_FILE = DATA_DIR / "state.json"
LOCK_FILE = DATA_DIR / "extract.lock"
LOG_FILE = DATA_DIR / "extract.log"
LOCK_STALE_SECONDS = 3 * 3600

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.yourcompany.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "25"))
SMTP_STARTTLS = os.environ.get("SMTP_STARTTLS", "false").lower() == "true"
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
MAIL_FROM = os.environ.get("MAIL_FROM", "lead-automation@yourcompany.com")
NOTIFY_TO = _env_list("NOTIFY_TO", "email-team@yourcompany.com,sv-team@yourcompany.com")
OWNER_EMAIL = os.environ.get("OWNER_EMAIL", "you@yourcompany.com")

DATA_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)


class AlreadyRunning(Exception):
    pass


# ------------------------------ LOCKING ------------------------------
@contextmanager
def single_instance():
    """Cross-platform lock using atomic file creation; stale locks expire."""
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        age = time.time() - LOCK_FILE.stat().st_mtime
        if age < LOCK_STALE_SECONDS:
            raise AlreadyRunning(f"Another run holds {LOCK_FILE} (age {int(age)}s)")
        logging.warning("Removing stale lock (age %ss)", int(age))
        LOCK_FILE.unlink()
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    try:
        os.write(fd, str(os.getpid()).encode())
        yield
    finally:
        os.close(fd)
        LOCK_FILE.unlink(missing_ok=True)


# ------------------------------ STATE ------------------------------
def load_state() -> dict:
    """Return {"watermark": ISO str, "seen": {lead_id: created_iso}}. Fails loudly if corrupt."""
    if not STATE_FILE.exists():
        logging.warning("No state file; first run starting from %s", INITIAL_SINCE)
        return {"watermark": INITIAL_SINCE, "seen": {}}
    try:
        state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        pd.Timestamp(state["watermark"])  # validate
        if not isinstance(state.setdefault("seen", {}), dict):
            raise ValueError("'seen' must be an object")
        return state
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(f"State file {STATE_FILE} is invalid: {exc}") from exc


def save_state(state: dict) -> None:
    """Atomic write: temp file, then replace."""
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    os.replace(tmp, STATE_FILE)


def to_utc(series: pd.Series) -> pd.Series:
    """Parse timestamps to UTC; naive values are interpreted in API_TIMEZONE."""
    ts = pd.to_datetime(series, errors="coerce")
    if ts.dt.tz is None:
        return ts.dt.tz_localize(API_TIMEZONE, ambiguous="NaT", nonexistent="NaT").dt.tz_convert("UTC")
    return ts.dt.tz_convert("UTC")


# ------------------------------ API ------------------------------
def build_session() -> requests.Session:
    # Retries 429/5xx and connection errors with exponential backoff and honours
    # Retry-After. 401/403 and other 4xx are NOT retried.
    retry = Retry(
        total=4,
        backoff_factor=2,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
        respect_retry_after_header=True,
    )
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    session.headers.update({"Authorization": f"Bearer {API_TOKEN}", "Accept": "application/json"})
    return session


def fetch_leads(session: requests.Session, since: str) -> list:
    """Fetch pages (oldest first) until an EMPTY page is returned."""
    url = f"{API_BASE}/{LEADS_ENDPOINT.lstrip('/')}"
    leads = []
    for page in range(1, MAX_PAGES + 1):
        resp = session.get(
            url,
            params={
                PARAM_SINCE: since,
                PARAM_SORT: SORT_VALUE,
                PARAM_PAGE: page,
                PARAM_PAGE_SIZE: PAGE_SIZE,
            },
            timeout=60,
        )
        resp.raise_for_status()
        try:
            batch = resp.json()[RESULTS_KEY]
        except (ValueError, KeyError, TypeError) as exc:
            raise ValueError(f"Unexpected API response format: {exc}") from exc
        if not isinstance(batch, list):
            raise ValueError(f"'{RESULTS_KEY}' is not a list")
        logging.info("Page %s: %s leads", page, len(batch))
        if not batch:  # empty page = done (safe even if the API caps page size)
            return leads
        leads.extend(batch)
    raise RuntimeError(f"Stopped after MAX_PAGES={MAX_PAGES}; possible runaway pagination")


# ------------------------------ EMAIL ------------------------------
def send_mail(subject: str, body: str, recipients: list) -> None:
    if not recipients:
        raise ValueError("No email recipients configured")
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = MAIL_FROM, ", ".join(recipients), subject
    msg.set_content(body)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        if SMTP_STARTTLS:
            smtp.starttls()
        if SMTP_USER:
            smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)


# ------------------------------ FILE ------------------------------
def write_excel(df: pd.DataFrame, run_time: pd.Timestamp) -> Path:
    """Write to a temp name, then rename, so readers never see a partial file."""
    SHARED_FOLDER.mkdir(parents=True, exist_ok=True)
    base = f"Leads_{run_time:%Y%m%d_%H%M%S}"
    final = SHARED_FOLDER / f"{base}.xlsx"
    counter = 1
    while final.exists():  # never overwrite an existing delivery
        final = SHARED_FOLDER / f"{base}_{counter}.xlsx"
        counter += 1
    part = final.with_name(final.name + ".part")
    df.to_excel(part, index=False, engine="openpyxl")
    os.replace(part, final)
    return final


# ------------------------------ MAIN ------------------------------
def run() -> None:
    run_time = pd.Timestamp.now(tz="UTC")
    state = load_state()
    watermark = pd.Timestamp(state["watermark"])
    if watermark.tzinfo is None:
        watermark = watermark.tz_localize("UTC")
    since = (watermark - OVERLAP).isoformat()
    logging.info("Run started. watermark=%s, querying since=%s", watermark, since)

    leads = fetch_leads(build_session(), since)
    if not leads:
        logging.info("No leads returned.")
        return

    df = pd.DataFrame(leads)
    missing = [c for c in (ID_FIELD, CREATED_FIELD) if c not in df.columns]
    if missing:
        raise ValueError(f"API records lack required field(s): {missing}")

    df["_created"] = to_utc(df[CREATED_FIELD])
    if df["_created"].isna().any():
        raise ValueError(f"Unparseable '{CREATED_FIELD}' values; refusing to move the checkpoint")

    # De-duplicate by lead ID within this pull (keep latest version of a record).
    before = len(df)
    df = df.drop_duplicates(subset=[ID_FIELD], keep="last")
    logging.info("Removed %s duplicate records within the pull", before - len(df))

    # Drop leads already delivered in earlier runs (overlap window protection).
    seen = state["seen"]
    new = df[~df[ID_FIELD].astype(str).isin(seen)].sort_values("_created")

    new_watermark = max(watermark, df["_created"].max())
    next_since = new_watermark - OVERLAP

    # Remember delivered IDs that could re-appear in the next overlap window.
    updated_seen = {k: v for k, v in seen.items() if pd.Timestamp(v) >= next_since}
    for lead_id, created in zip(new[ID_FIELD].astype(str), new["_created"]):
        if created >= next_since:
            updated_seen[lead_id] = created.isoformat()

    if new.empty:
        logging.info("Only previously delivered leads returned; no file created.")
        save_state({"watermark": new_watermark.isoformat(), "seen": updated_seen})
        return

    file_path = write_excel(new.drop(columns="_created"), run_time)
    logging.info("Saved %s leads to %s", len(new), file_path)

    # Checkpoint moves right after the file is safely saved, so a mail failure
    # can never cause the same leads to be extracted into a second file.
    save_state({"watermark": new_watermark.isoformat(), "seen": updated_seen})

    try:
        send_mail(
            subject=f"Lead extract ready - {len(new)} leads ({run_time:%d-%b %H:%M UTC})",
            body=(
                "The latest lead extract is ready.\n\n"
                f"File: {file_path}\n"
                f"Leads: {len(new)}\n"
                f"Generated: {run_time:%Y-%m-%d %H:%M:%S UTC}"
            ),
            recipients=NOTIFY_TO,
        )
    except Exception as exc:
        logging.exception("File saved but notification failed")
        raise RuntimeError(f"File saved at {file_path} but team notification failed: {exc}") from exc

    logging.info("Run completed successfully.")


def main() -> None:
    try:
        with single_instance():
            run()
    except AlreadyRunning as exc:
        logging.warning("Skipped: %s", exc)
    except Exception as exc:
        logging.exception("Run failed")
        try:
            send_mail("Lead extract FAILED", f"The Haloocom extract failed:\n\n{exc}", [OWNER_EMAIL])
        except Exception:
            logging.exception("Could not send failure alert")
        raise


if __name__ == "__main__":
    main()
