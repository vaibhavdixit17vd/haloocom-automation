"""
Haloocom Prospect Detail Report Extraction Automation.

Handles CSV payload response from prospect_detail_report.php.
"""

import io
import json
import logging
import os
import smtplib
import time
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def _env_list(name: str, default: str = "") -> list:
    return [x.strip() for x in os.environ.get(name, default).split(",") if x.strip()]


def _get_default_data_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


# ------------------------------ CONFIG ------------------------------
API_ENDPOINT = os.environ.get(
    "HALOOCOM_API_ENDPOINT",
    "https://zdpm.haloocom.in/haloocomConnect5/prospect_detail_report.php",
)

ID_FIELD = os.environ.get("HALOOCOM_ID_FIELD", "Lead ID")
INITIAL_LOOKBACK_DAYS = int(os.environ.get("INITIAL_LOOKBACK_DAYS", "1"))
OVERLAP_DAYS = int(os.environ.get("OVERLAP_DAYS", "1"))

SHARED_FOLDER = Path(os.environ.get("LEADS_SHARED_FOLDER", r"\\fs05\mis2\Vaibhav Dixit\Haloocom Downloads"))
DATA_DIR = Path(os.environ.get("LEADS_DATA_DIR", _get_default_data_dir()))
STATE_FILE = DATA_DIR / "state.json"
LOCK_FILE = DATA_DIR / "extract.lock"
LOG_FILE = DATA_DIR / "extract.log"
LOCK_STALE_SECONDS = 3600

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
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        age = time.time() - LOCK_FILE.stat().st_mtime
        if age < LOCK_STALE_SECONDS:
            raise AlreadyRunning(f"Another run holds {LOCK_FILE} (age {int(age)}s)")
        logging.warning("Removing stale lock (age %ss)", int(age))
        LOCK_FILE.unlink(missing_ok=True)
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    try:
        os.write(fd, str(os.getpid()).encode())
        yield
    finally:
        os.close(fd)
        LOCK_FILE.unlink(missing_ok=True)


# ------------------------------ STATE ------------------------------
def load_state() -> dict:
    default_start = (datetime.now(timezone.utc) - timedelta(days=INITIAL_LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    if not STATE_FILE.exists():
        logging.warning("No state file found; first run starting from %s", default_start)
        return {"watermark_date": default_start, "seen": []}
    try:
        state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        datetime.strptime(state["watermark_date"], "%Y-%m-%d")
        if not isinstance(state.setdefault("seen", []), list):
            state["seen"] = []
        return state
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(f"State file {STATE_FILE} is invalid: {exc}") from exc


def save_state(state: dict) -> None:
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    os.replace(tmp, STATE_FILE)


# ------------------------------ API FETCH ------------------------------
def build_session() -> requests.Session:
    retry = Retry(
        total=4,
        backoff_factor=2,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET",),
    )
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    return session


def fetch_prospect_report(session: requests.Session, start_date: str, end_date: str) -> pd.DataFrame:
    params = {"start_date": start_date, "end_date": end_date}
    logging.info("Requesting report from %s with params: %s", API_ENDPOINT, params)

    resp = session.get(API_ENDPOINT, params=params, timeout=120)
    resp.raise_for_status()

    raw_text = resp.text.strip()

    if not raw_text:
        return pd.DataFrame()

    # Clean Haloocom prefix header text if present
    prefix = "Prospect Detail Report Downloaded Successfully"
    if raw_text.startswith(prefix):
        raw_text = raw_text[len(prefix):].lstrip()

    try:
        df = pd.read_csv(io.StringIO(raw_text))
        return df
    except Exception as exc:
        raise ValueError(f"Failed to parse CSV response: {exc}. Raw snippet: {raw_text[:200]}") from exc


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


# ------------------------------ FILE OUTPUT ------------------------------
def write_excel(df: pd.DataFrame, run_time: datetime) -> Path:
    SHARED_FOLDER.mkdir(parents=True, exist_ok=True)
    base = f"Haloocom_Leads_{run_time:%Y%m%d_%H%M%S}"
    final = SHARED_FOLDER / f"{base}.xlsx"
    counter = 1
    while final.exists():
        final = SHARED_FOLDER / f"{base}_{counter}.xlsx"
        counter += 1
    part = final.with_name(final.name + ".part")
    df.to_excel(part, index=False, engine="openpyxl")
    os.replace(part, final)
    return final


# ------------------------------ MAIN WORKFLOW ------------------------------
def run() -> None:
    now_utc = datetime.now(timezone.utc)
    today_str = now_utc.strftime("%Y-%m-%d")

    state = load_state()
    watermark_dt = datetime.strptime(state["watermark_date"], "%Y-%m-%d")

    start_dt = max(watermark_dt - timedelta(days=OVERLAP_DAYS), datetime(2020, 1, 1))
    start_date_str = start_dt.strftime("%Y-%m-%d")
    end_date_str = today_str

    logging.info("Starting run. Date query range: %s to %s", start_date_str, end_date_str)

    df = fetch_prospect_report(build_session(), start_date_str, end_date_str)
    if df.empty:
        logging.info("No prospect records returned by API.")
        save_state({"watermark_date": today_str, "seen": state.get("seen", [])})
        return

    # Check for Lead ID column
    if ID_FIELD in df.columns:
        effective_id_field = ID_FIELD
    elif "SNO" in df.columns:
        effective_id_field = "SNO"
    else:
        logging.warning("ID field not found in columns (%s). Creating composite hash ID.", list(df.columns))
        df["_extracted_id"] = df.astype(str).apply(lambda row: hash("".join(row)), axis=1).astype(str)
        effective_id_field = "_extracted_id"

    # Deduplicate within current pull
    df = df.drop_duplicates(subset=[effective_id_field], keep="last")

    # Filter out previously delivered IDs
    seen_set = set(state.get("seen", []))
    new_df = df[~df[effective_id_field].astype(str).isin(seen_set)].copy()

    if "_extracted_id" in new_df.columns:
        new_df = new_df.drop(columns=["_extracted_id"])

    if new_df.empty:
        logging.info("No new prospect records found after de-duplication.")
        save_state({"watermark_date": today_str, "seen": list(seen_set)})
        return

    file_path = write_excel(new_df, now_utc)
    logging.info("Saved %s new leads to %s", len(new_df), file_path)

    # Save state
    updated_seen = list(seen_set.union(set(df[effective_id_field].astype(str))))[-10000:]
    save_state({"watermark_date": today_str, "seen": updated_seen})

    # Email notification
    try:
        send_mail(
            subject=f"Haloocom Leads Extract Ready - {len(new_df)} leads ({now_utc:%d-%b %H:%M UTC})",
            body=(
                "The latest prospect detail report extract is ready.\n\n"
                f"File: {file_path}\n"
                f"New Leads Delivered: {len(new_df)}\n"
                f"Generated: {now_utc:%Y-%m-%d %H:%M:%S UTC}"
            ),
            recipients=NOTIFY_TO,
        )
    except Exception as exc:
        logging.exception("File written, but notification email failed")
        raise RuntimeError(f"File saved to {file_path}, but email failed: {exc}") from exc

    logging.info("Run finished successfully.")


def main() -> None:
    try:
        with single_instance():
            run()
    except AlreadyRunning as exc:
        logging.warning("Skipped run: %s", exc)
    except Exception as exc:
        logging.exception("Run failed")
        try:
            send_mail("Haloocom Extract FAILED", f"The extraction script encountered an error:\n\n{exc}", [OWNER_EMAIL])
        except Exception:
            logging.exception("Could not send failure alert email")
        raise


if __name__ == "__main__":
    main()