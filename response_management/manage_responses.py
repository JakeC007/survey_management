"""
manage_responses.py
───────────────────
Post-survey response management pipeline for the Teen AI Survey.

How you use it
──────────────
  1. Download the Qualtrics RESPONSE SURVEY export ZIP (the teen survey,
     not the consent/screener survey).
  2. Drop it into the shared ingest/ folder (survey_management/ingest/).
  3. Double-click one of the runner .bat files in this folder, or run:

         ..\\.venv\\Scripts\\python manage_responses.py

     The script will:
       • Validate the ZIP isn't a consent survey file (wrong-ZIP protection)
       • Match completed responses to invited participants via the cid field
       • Update the master tracker with completion timestamps
       • (Gift-card / payment handling now lives in payment_management/)
       • Send reminder emails to invitees who haven't completed yet,
         based on configured follow-up intervals (3 days → follow-up 1,
         5 days → follow-up 2)
       • Ignore all entries dated before the configured cutoff date

Safety flags (optional)
───────────────────────
  --draft     Stage reminder emails as Outlook Drafts instead of sending.
  --dry-run   Update the tracker only. No Outlook, no emails, no deletions.

Outputs (all inside survey_management/data/)
─────────────────────────────────────────────
  data/participant_tracker_auto.xlsx  — Updates survey_completed_at,
                                        follow_up_1_sent_at, follow_up_2_sent_at.
  data/send_log.csv                   — Appends reminder email records.
  data/reminder_log.csv               — Append-only reminder audit log.
  (console)                           — Gift card list printed after each run.

Configuration
─────────────
  All settings are read from ../config.yaml (the shared config file in
  survey_management/).  See README.md for the full config-key mapping.

Wrong-ZIP protection
────────────────────
  Because this script shares the ingest/ folder with send_survey_emails.py,
  you might accidentally drop a consent survey ZIP instead of a response
  survey ZIP.  The script inspects the CSV headers inside every ZIP before
  processing.  If consent-survey-specific columns are detected, it logs a
  clear warning and exits cleanly without touching any data.

Date cutoff
───────────
  All consent entries and response survey rows whose Recorded Date is
  strictly before response_cutoff_date (default 2026-06-09) are silently
  skipped.  This prevents stale pre-pilot data from affecting tracking.
"""

import argparse
import io
import sys
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import yaml


# ── CONFIG & PATH SETUP ────────────────────────────────────────────────────────
#
# This script lives in survey_management/response_management/.
# The shared config.yaml and all bookkeeping files live one level up, in
# survey_management/.  All paths are resolved from REPO_ROOT so this script
# reads/writes the exact same files as send_survey_emails.py.

SCRIPT_DIR = Path(__file__).parent          # …/survey_management/response_management/
REPO_ROOT  = SCRIPT_DIR.parent              # …/survey_management/
CONFIG_PATH = REPO_ROOT / "config.yaml"

with open(CONFIG_PATH, "r", encoding="utf-8") as _f:
    config = yaml.safe_load(_f)

# ── GENERAL SETTINGS ──────────────────────────────────────────────────────────
# config key → variable
# "survey_link"                → SURVEY_LINK  (base URL; cid appended per participant)
# "sender.name"                → SENDER_NAME
# "sender.title"               → SENDER_TITLE
# "sender.contact_email"       → CONTACT_EMAIL
# "shared_mailbox"             → SHARED_MAILBOX

SURVEY_LINK    = config["survey_link"]
SENDER_NAME    = config["sender"]["name"]
SENDER_TITLE   = config["sender"]["title"]
CONTACT_EMAIL  = config["sender"]["contact_email"]
SHARED_MAILBOX = config["shared_mailbox"]

# ── PATHS ─────────────────────────────────────────────────────────────────────
# All paths resolve from REPO_ROOT (survey_management/) so both scripts share
# the same files.
#
# config key                   → variable
# "paths.inbox_dir"            → INBOX_DIR         (shared ingest folder)
# "paths.tracker_file"         → TRACKER_PATH      (shared XLSX tracker)
# "paths.send_log_file"        → SEND_LOG_PATH     (shared send log CSV)
# "paths.reminder_log_file"    → REMINDER_LOG_PATH (reminder audit CSV)

INBOX_DIR         = REPO_ROOT / config["paths"]["inbox_dir"]
TRACKER_PATH      = REPO_ROOT / config["paths"]["tracker_file"]
SEND_LOG_PATH     = REPO_ROOT / config["paths"]["send_log_file"]
REMINDER_LOG_PATH = REPO_ROOT / config["paths"].get("reminder_log_file",
                                                      "data/reminder_log.csv")

# ── DATE CUTOFF ───────────────────────────────────────────────────────────────
# config key: "response_cutoff_date"
# Entries (consent or response) recorded before this date are ignored.
CUTOFF_DATE = date.fromisoformat(config.get("response_cutoff_date", "2026-06-09"))

# ── REMINDER INTERVALS ────────────────────────────────────────────────────────
# config keys: "reminder_intervals.follow_up_1_days"
#              "reminder_intervals.follow_up_2_days"
_ri = config.get("reminder_intervals", {})
FOLLOW_UP_1_DAYS = int(_ri.get("follow_up_1_days", 3))
FOLLOW_UP_2_DAYS = int(_ri.get("follow_up_2_days", 5))

# ── RESPONSE SURVEY FIELD NAMES ───────────────────────────────────────────────
# config keys: "response_qualtrics_fields.*"
# These match columns in the RESPONSE SURVEY export (not the consent export).
#
# "response_qualtrics_fields.response_id"   → RF_RESPONSE_ID
# "response_qualtrics_fields.recorded_date" → RF_RECORDED_DATE
# "response_qualtrics_fields.finished"      → RF_FINISHED
# "response_qualtrics_fields.cid"           → RF_CID
#   cid is the embedded-data field capturing the ?cid= URL parameter that
#   send_survey_emails.py appends to each invitation link.  It equals the
#   participant's consent Response ID, so it links completions back to the
#   tracker rows.

_rqf = config.get("response_qualtrics_fields", {})
RF_RESPONSE_ID   = _rqf.get("response_id",   "Response ID")
RF_RECORDED_DATE = _rqf.get("recorded_date",  "Recorded Date")
RF_FINISHED      = _rqf.get("finished",       "Finished")
RF_CID           = _rqf.get("cid",            "cid")

# ── CONSENT-ZIP DETECTION MARKERS ─────────────────────────────────────────────
# These normalised substrings appear in consent survey headers and NOT in
# response survey headers.  Any match means the wrong ZIP was dropped.
CONSENT_HEADER_MARKERS = [
    "parent/guardian full name",
    "child participant's full name",
    "i confirm that i have read",
    "i understand what the researchers asked",
]


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _norm(s) -> str:
    """Normalise a label: lowercase, smart-quote/dash unification, collapse whitespace."""
    s = "" if s is None else str(s)
    for a, b in (("'", "'"), ("'", "'"), ("“", '"'), ("”", '"'),
                 ("–", "-"), ("—", "-")):
        s = s.replace(a, b)
    return " ".join(s.lower().split())


def load_csv_from_zip(zip_path: str) -> pd.DataFrame:
    """Extract the first CSV from a Qualtrics export ZIP and return data rows
    with the human-readable label row as the column header.

    Handles the standard three-row Qualtrics export format (short var names,
    human-readable labels, ImportId JSON) as well as simpler two-row exports.
    """
    with zipfile.ZipFile(zip_path) as z:
        csv_names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            raise ValueError("No CSV found inside the ZIP.")
        raw = z.read(csv_names[0])

    grid = pd.read_csv(io.BytesIO(raw), dtype=str, header=None, encoding="utf-8-sig")

    # Locate the label row: the first of the top 3 rows containing "Response ID".
    label_row = 0
    for i in range(min(3, len(grid))):
        if grid.iloc[i].map(lambda v: _norm(v) == "response id").any():
            label_row = i
            break

    grid.columns = grid.iloc[label_row]
    data = grid.iloc[label_row + 1:]
    # Drop the ImportId metadata row if present directly after the label row.
    if len(data) and data.iloc[0].map(lambda v: str(v).startswith('{"ImportId"')).any():
        data = data.iloc[1:]
    return data.reset_index(drop=True).fillna("")


def is_consent_zip(df: pd.DataFrame) -> bool:
    """Return True if this DataFrame looks like a consent survey export.

    Checks column headers for consent-specific substrings.  If any match,
    this is the wrong ZIP and processing should be aborted.
    """
    norm_cols = [_norm(c) for c in df.columns]
    return any(
        any(marker in col for col in norm_cols)
        for marker in CONSENT_HEADER_MARKERS
    )


def find_column(df: pd.DataFrame, needle: str):
    """Return the first column whose normalised name contains `needle`, or None."""
    for col in df.columns:
        if needle in _norm(col):
            return col
    return None


def find_column_exact(df: pd.DataFrame, *names: str):
    """Return the first column whose normalised name EXACTLY equals one of `names`.

    Use this for short, ambiguous field names like `cid` where a loose substring
    match would wrongly grab a longer column.  For example the survey item
    "I feel comfortable deciding what information I should share with an AI tool"
    contains the substring "cid" (in "deciding"), and the flag column "cid_f"
    contains "cid" as a prefix — both would be matched by find_column(df, "cid"),
    causing the wrong column (answer text) to be read as the participant cid.
    """
    targets = {_norm(n) for n in names}
    for col in df.columns:
        if _norm(col) in targets:
            return col
    return None


def _peek_header_text(path) -> str:
    """Lowercased text of the first 3 (header) rows of the export's first CSV."""
    with zipfile.ZipFile(path) as z:
        csv_names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            return ""
        raw = z.read(csv_names[0])
    grid = pd.read_csv(io.BytesIO(raw), dtype=str, header=None,
                       encoding="utf-8-sig", nrows=3)
    return " || ".join(_norm(v) for v in grid.values.flatten())


def pick_response_zip(zips):
    """Most recently MODIFIED ZIP whose headers are NOT a consent export — i.e.
    the teen response survey. The inbox holds both types; consent ZIPs and
    unreadable ZIPs are skipped. Cheap: only the header rows are read."""
    for zp in sorted(zips, key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            text = _peek_header_text(zp)
        except Exception:
            continue
        if not text.strip():
            continue
        if not any(m in text for m in CONSENT_HEADER_MARKERS):
            return zp
    return None


def parse_recorded_date(date_str: str):
    """Parse a Qualtrics 'Recorded Date' string to a date object, or None on failure."""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
                "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


def parse_emailed_at(ts_str: str):
    """Parse an emailed_at ISO timestamp to a date object, or None."""
    if not ts_str or not ts_str.strip():
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(ts_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


# ── TRACKER I/O ───────────────────────────────────────────────────────────────

def load_tracker() -> pd.DataFrame:
    """Load the Participants sheet from the tracker XLSX.  Returns an empty
    DataFrame with the required columns if the file doesn't exist yet."""
    required_cols = config.get("tracker_columns", [])
    if not TRACKER_PATH.exists():
        return pd.DataFrame(columns=required_cols)
    return pd.read_excel(TRACKER_PATH, sheet_name="Participants", dtype=str).fillna("")


def save_tracker(df: pd.DataFrame):
    """Write the updated Participants DataFrame back to the tracker XLSX,
    preserving the Run Stats sheet from the previous run."""
    TRACKER_PATH.parent.mkdir(parents=True, exist_ok=True)
    # Load existing Run Stats if available; otherwise create a stub.
    if TRACKER_PATH.exists():
        try:
            stats_df = pd.read_excel(TRACKER_PATH, sheet_name="Run Stats",
                                     dtype=str).fillna("")
        except Exception:
            stats_df = pd.DataFrame(columns=["metric", "count"])
    else:
        stats_df = pd.DataFrame(columns=["metric", "count"])

    with pd.ExcelWriter(TRACKER_PATH, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Participants", index=False)
        stats_df.to_excel(xl, sheet_name="Run Stats", index=False)


# ── LOGGING ───────────────────────────────────────────────────────────────────

def append_send_log(rows: list):
    """Append reminder rows to the shared send_log.csv."""
    if not rows:
        return
    SEND_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = pd.DataFrame(rows, columns=["response_id", "child_name", "delivery_email",
                                      "status", "mode", "emailed_at"])
    out.to_csv(SEND_LOG_PATH, mode="a", header=not SEND_LOG_PATH.exists(),
               index=False, encoding="utf-8")


def append_reminder_log(rows: list):
    """Append rows to the reminder-specific audit log."""
    if not rows:
        return
    REMINDER_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = pd.DataFrame(rows, columns=["response_id", "child_name", "delivery_email",
                                      "reminder_type", "sent_at", "mode",
                                      "days_since_invite"])
    out.to_csv(REMINDER_LOG_PATH, mode="a", header=not REMINDER_LOG_PATH.exists(),
               index=False, encoding="utf-8")


# ── EMAIL TEMPLATES ───────────────────────────────────────────────────────────

def build_reminder_body(child_first: str, follow_up_n: int,
                        unique_url: str, sender_name: str,
                        sender_title: str, contact: str) -> str:
    """Build the HTML body for a follow-up reminder email."""
    if follow_up_n == 1:
        opener = (
            f"<p>Hi {child_first},</p>"
            f"<p>Just a friendly reminder — a few days ago we invited you to take "
            f"a short survey about AI and privacy in schools. We haven't seen your "
            f"response yet, and we'd love to hear from you!</p>"
        )
    else:
        opener = (
            f"<p>Hi {child_first},</p>"
            f"<p>This is our final reminder about the survey on AI and privacy in "
            f"schools. We still haven't received your response. It takes only "
            f"10–15 minutes, and your perspective really does matter.</p>"
        )

    return f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
{opener}

<p style="margin:20px 0;">
  <a href="{unique_url}" style="background:#800000;color:#fff;padding:10px 20px;
  text-decoration:none;border-radius:4px;font-weight:bold;">Take the Survey →</a>
</p>

<p>Or copy this link into your browser:<br>
<a href="{unique_url}">{unique_url}</a></p>

<p><strong>Reminder:</strong> Your parent or guardian already gave permission for
you to participate. Completing the survey earns you a
<strong>$10 Amazon gift card</strong>.</p>

<p>Thank you — we really do want to hear from you.</p>

<p>Best regards,<br>
<strong>{sender_name}</strong><br>
{sender_title}<br>
Investigations of Educational Technology to Safeguard Children's Privacy<br>
University of Chicago<br>
<a href="mailto:{contact}">{contact}</a></p>
</body></html>"""


# ── OUTLOOK ───────────────────────────────────────────────────────────────────

_OUTLOOK: dict = {}


def _get_outlook():
    import win32com.client as win32
    if "app" not in _OUTLOOK:
        _OUTLOOK["app"] = win32.Dispatch("outlook.application")
    return _OUTLOOK["app"]


def _shared_drafts_folder():
    """Return the shared mailbox's Drafts folder, or None if not configured."""
    if not SHARED_MAILBOX:
        return None
    if "drafts" in _OUTLOOK:
        return _OUTLOOK["drafts"]
    ns = _get_outlook().GetNamespace("MAPI")
    recip = ns.CreateRecipient(SHARED_MAILBOX)
    recip.Resolve()
    if not recip.Resolved:
        raise RuntimeError(
            f"Outlook could not resolve shared mailbox '{SHARED_MAILBOX}'. "
            "Check the address and that the mailbox is added to your Outlook profile.")
    OL_FOLDER_DRAFTS = 16
    folder = ns.GetSharedDefaultFolder(recip, OL_FOLDER_DRAFTS)
    _OUTLOOK["drafts"] = folder
    return folder


def _assert_writable(path: Path, label: str):
    """Raise a clear error if `path` can't be written (typically because it's
    open in Excel, which takes an exclusive lock on Windows). Writes nothing."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if path.exists():
            with open(path, "a"):   # append-open touches no bytes, just checks the lock
                pass
        else:
            probe = path.with_name(path.name + ".writetest")
            with open(probe, "w"):
                pass
            probe.unlink()
    except Exception as e:
        raise RuntimeError(
            f"Cannot write {label} ({path}). It is most likely open in Excel — "
            f"close it and run again. Original error: {e}")


def preflight(draft_only: bool, dry_run: bool):
    """Fail fast BEFORE touching any recipient.

    1. The output files (tracker, send_log.csv, reminder_log.csv) are writable.
       Critical: every mode writes these at the end, so a reminder must never go
       out only to fail to record itself. Checked in all modes.
    2. Outlook is reachable (skipped in dry-run, which sends nothing)."""
    _assert_writable(TRACKER_PATH, "the participant tracker")
    _assert_writable(SEND_LOG_PATH, "send_log.csv")
    _assert_writable(REMINDER_LOG_PATH, "reminder_log.csv")
    if dry_run:
        return
    try:
        _get_outlook()
    except Exception as e:
        raise RuntimeError(
            "Could not connect to Outlook. Make sure classic Outlook (NOT New "
            "Outlook) is open and signed in. Details: " + str(e))
    if SHARED_MAILBOX and draft_only:
        _shared_drafts_folder()


def send_via_outlook(to_email: str, subject: str, html_body: str, draft_only: bool):
    mail = _get_outlook().CreateItem(0)   # olMailItem
    mail.To = to_email
    mail.Subject = subject
    mail.HTMLBody = html_body
    if SHARED_MAILBOX:
        mail.SentOnBehalfOfName = SHARED_MAILBOX
    if draft_only:
        mail.Save()
        folder = _shared_drafts_folder()
        if folder is not None:
            mail.Move(folder)
    else:
        mail.Send()


# ── RESPONSE ZIP PROCESSING ───────────────────────────────────────────────────

def process_response_zips(zips: list, tracker_df: pd.DataFrame
                          ) -> tuple[pd.DataFrame, list]:
    """Ingest response survey ZIPs and update tracker completion status.

    Returns the (possibly modified) tracker DataFrame and a list of flagged
    issues (e.g. missing cid values).
    """
    flags = []

    for zp in zips:
        print(f"── Ingesting response ZIP: {zp.name} ──")

        try:
            df = load_csv_from_zip(str(zp))
        except Exception as e:
            print(f"  ERROR  Could not read {zp.name}: {e}  (skipping)\n")
            continue

        # ── Wrong-ZIP guard ──────────────────────────────────────────────────
        if is_consent_zip(df):
            print(
                f"\n  WARNING: Detected a Consent Survey ZIP instead of a "
                f"Response Survey ZIP.\n"
                f"  File: {zp.name}\n"
                f"  This file looks like a screener/consent export "
                f"(it contains consent-specific columns).\n"
                f"  Skipping — no data was changed.\n"
                f"  To process consent data, use send_survey_emails.py instead.\n"
            )
            continue

        # ── Locate required columns ──────────────────────────────────────────
        col_rid  = find_column(df, "response id")
        col_date = find_column(df, "recorded date")
        col_done = find_column(df, "finished")
        # cid must match EXACTLY — a loose substring match grabs "cid_f" or any
        # question label containing "cid" (e.g. "...comfortable deciding...").
        col_cid  = find_column_exact(df, RF_CID)

        if col_cid is None:
            # cid column not found at all — probably wrong survey or cid not set up
            print(
                f"  WARNING: No '{RF_CID}' column found in {zp.name}.\n"
                f"  Check that 'cid' is configured as embedded data in your "
                f"Qualtrics response survey flow.\n"
                f"  All rows in this file will be flagged as missing-cid.\n"
            )

        rows_processed = 0
        rows_completed = 0
        rows_missing_cid = 0
        rows_before_cutoff = 0

        for _, row in df.iterrows():
            resp_id = str(row.get(col_rid, "")).strip() if col_rid else ""
            rec_date_str = str(row.get(col_date, "")).strip() if col_date else ""
            finished = (str(row.get(col_done, "")).strip().upper()
                        in ("TRUE", "1")) if col_done else False
            cid = str(row.get(col_cid, "")).strip() if col_cid else ""

            rows_processed += 1

            # ── Date cutoff ─────────────────────────────────────────────────
            rec_date = parse_recorded_date(rec_date_str) if rec_date_str else None
            if rec_date is not None and rec_date < CUTOFF_DATE:
                rows_before_cutoff += 1
                continue

            # ── Missing cid ─────────────────────────────────────────────────
            if not cid:
                rows_missing_cid += 1
                flags.append({
                    "zip": zp.name,
                    "response_id": resp_id,
                    "recorded_date": rec_date_str,
                    "issue": "missing_cid",
                })
                print(f"  FLAG  Response {resp_id or '(no id)'} on "
                      f"{rec_date_str or 'unknown date'} has no cid — cannot match.")
                continue

            if not finished:
                continue   # not completed — nothing to update

            # ── Match to tracker ────────────────────────────────────────────
            mask = tracker_df["response_id"] == cid
            if not mask.any():
                flags.append({
                    "zip": zp.name,
                    "response_id": resp_id,
                    "recorded_date": rec_date_str,
                    "issue": f"cid_not_in_tracker: cid={cid}",
                })
                print(f"  FLAG  cid={cid} not found in tracker "
                      f"(response {resp_id}) — unmatched completion.")
                continue

            # Only stamp completion once (first completion wins)
            already_done = str(
                tracker_df.loc[mask, "survey_completed_at"].values[0]
            ).strip()
            if not already_done:
                completed_ts = (rec_date_str if rec_date_str
                                else datetime.now().isoformat(timespec="seconds"))
                tracker_df.loc[mask, "survey_completed_at"] = completed_ts
                rows_completed += 1
                child = str(tracker_df.loc[mask, "child_name"].values[0]).strip()
                print(f"  COMPLETED  {child or cid} (cid={cid})")

        print(f"  Rows: {rows_processed} total | "
              f"{rows_before_cutoff} before cutoff (skipped) | "
              f"{rows_missing_cid} missing cid | {rows_completed} newly marked complete\n")

    return tracker_df, flags


# ── PAYMENT / GIFT CARDS ────────────────────────────────────────────────────────
# Gift-card / payment logic now lives in payment_management/manage_payments.py.
# This script only tracks completions and sends reminders; it no longer prints a
# gift card list. Run payment_management to see who is owed a card and to manage
# the paid/unpaid ledger.


# ── REMINDER LOGIC ─────────────────────────────────────────────────────────────

def check_and_send_reminders(tracker_df: pd.DataFrame,
                             draft_only: bool,
                             dry_run: bool,
                             now_iso: str) -> tuple[pd.DataFrame, list, list]:
    """Identify participants due for follow-up reminders and send/draft them.

    Timing model (spacing-based, not t=0-based):
      - Follow-up 1: FOLLOW_UP_1_DAYS days after the original invitation.
      - Follow-up 2: FOLLOW_UP_2_DAYS days after follow-up 1 was sent.
        If follow-up 1 hasn't been sent yet there is no anchor, so follow-up 2
        is skipped until follow-up 1 goes out (possibly in the same run).

    This means a late first run doesn't collapse the gap between follow-ups —
    each reminder is spaced off the previous one.

    Returns the (updated) tracker DataFrame, send_log rows, and reminder_log rows.
    """
    today = date.today()
    send_log_rows = []
    reminder_log_rows = []
    mode_label = "DRY-RUN" if dry_run else ("DRAFT" if draft_only else "SEND")
    counts = Counter()

    for idx, row in tracker_df.iterrows():
        emailed_at_str = str(row.get("emailed_at", "")).strip()
        if not emailed_at_str:
            continue   # never invited — skip

        invite_date = parse_emailed_at(emailed_at_str)
        if invite_date is None:
            continue

        # Date-cutoff filter
        if invite_date < CUTOFF_DATE:
            continue

        # Already completed — no reminders needed
        if str(row.get("survey_completed_at", "")).strip():
            continue

        rid    = str(row.get("response_id", "")).strip()
        child  = str(row.get("child_name", "")).strip()
        email  = str(row.get("delivery_email", "")).strip()

        if not email or "@" not in email:
            continue   # no valid email

        child_first = child.split()[0] if child.strip() else child
        unique_url  = f"{SURVEY_LINK}?cid={rid}"

        for follow_up_n, threshold, col in [
            (1, FOLLOW_UP_1_DAYS, "follow_up_1_sent_at"),
            (2, FOLLOW_UP_2_DAYS, "follow_up_2_sent_at"),
        ]:
            if str(row.get(col, "")).strip():
                continue   # already sent this follow-up

            # Determine the anchor date for this follow-up's spacing.
            # Follow-up 1 is anchored to the original invitation.
            # Follow-up 2 is anchored to when follow-up 1 was sent — if
            # follow-up 1 hasn't gone out yet, there's no anchor so we skip.
            if follow_up_n == 1:
                anchor_date = invite_date
            else:
                fu1_sent_str = str(row.get("follow_up_1_sent_at", "")).strip()
                if not fu1_sent_str:
                    continue  # follow-up 1 not sent yet; no anchor for follow-up 2
                anchor_date = parse_emailed_at(fu1_sent_str)
                if anchor_date is None:
                    continue

            days_since_anchor = (today - anchor_date).days
            if days_since_anchor < threshold:
                continue

            status_key = f"REMINDER_{follow_up_n}"
            subject = (f"Reminder: Share Your Thoughts on AI in School "
                       f"& Earn $10 (Follow-up {follow_up_n})")
            body = build_reminder_body(
                child_first, follow_up_n, unique_url,
                SENDER_NAME, SENDER_TITLE, CONTACT_EMAIL,
            )

            anchor_label = "invite" if follow_up_n == 1 else "follow-up 1"
            if dry_run:
                sent_at = ""
                status  = f"ELIGIBLE (dry-run) — {status_key}"
                print(f"  WOULD REMIND ({follow_up_n})  {child_first} → {email}  "
                      f"({days_since_anchor}d since {anchor_label})")
            else:
                try:
                    send_via_outlook(email, subject, body, draft_only)
                    sent_at = datetime.now().isoformat(timespec="seconds")
                    status  = f"DRAFTED_{follow_up_n}" if draft_only else f"SENT_{follow_up_n}"
                    tracker_df.at[idx, col] = sent_at
                    print(f"  {status}  {child_first} → {email}  "
                          f"({days_since_anchor}d since {anchor_label})")
                except Exception as e:
                    sent_at = ""
                    status  = f"ERROR_{follow_up_n}"
                    print(f"  ERROR ({follow_up_n})  {child_first} → {email}: {e}")

            counts[status] += 1
            send_log_rows.append({
                "response_id":    rid,
                "child_name":     child,
                "delivery_email": email,
                "status":         status,
                "mode":           mode_label,
                "emailed_at":     sent_at,
            })
            reminder_log_rows.append({
                "response_id":      rid,
                "child_name":       child,
                "delivery_email":   email,
                "reminder_type":    f"follow_up_{follow_up_n}",
                "sent_at":          sent_at,
                "mode":             mode_label,
                "days_since_anchor": days_since_anchor,
            })

    return tracker_df, send_log_rows, reminder_log_rows, counts


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Post-survey response management: match completions, generate gift "
            "card list, send reminder emails.  Drop the response survey ZIP in "
            "the ingest/ folder before running."
        )
    )
    parser.add_argument("--draft", action="store_true",
                        help="Stage reminder emails as Outlook Drafts instead of sending.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Update tracker only — no Outlook, no emails.")
    args = parser.parse_args()

    dry_run    = args.dry_run
    draft_only = args.draft
    mode_label = "DRY-RUN" if dry_run else ("DRAFT" if draft_only else "SEND")

    print(f"\nMode       : {mode_label}")
    print(f"Inbox      : {INBOX_DIR}")
    print(f"Tracker    : {TRACKER_PATH}")
    print(f"Cutoff date: {CUTOFF_DATE}  (entries before this date are ignored)")
    print(f"Follow-up 1: {FOLLOW_UP_1_DAYS} days after invite")
    print(f"Follow-up 2: {FOLLOW_UP_2_DAYS} days after invite\n")

    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    # The inbox holds both consent and response exports and is never emptied.
    # Pick the most recently MODIFIED ZIP whose headers match the response-survey
    # schema; quietly skip consent ZIPs. Completion stamping is idempotent
    # (first completion wins), so re-reading an export changes nothing.
    _all_zips = sorted(INBOX_DIR.glob("*.zip"), key=lambda p: p.stat().st_mtime)
    _chosen = pick_response_zip(_all_zips)
    zips = [_chosen] if _chosen else []
    if _all_zips and not zips:
        print(f"Found {len(_all_zips)} ZIP(s) in ingest but none match the "
              f"response-survey schema (they look like consent exports).\n")
    elif zips:
        _others = [z.name for z in _all_zips if z != _chosen]
        print(f"Using response export: {_chosen.name}")
        if _others:
            print(f"  (ignoring {len(_others)} other ZIP(s): {', '.join(_others)})")
        print()

    # ── Fail fast: output files must be writable (all modes write them), and
    #    Outlook must be reachable when we'll actually send. ────────────────────
    try:
        preflight(draft_only, dry_run)
    except Exception as e:
        print(f"PREFLIGHT FAILED: {e}\n")
        return

    # ── Load tracker ──────────────────────────────────────────────────────────
    if not TRACKER_PATH.exists():
        print(
            f"WARNING: Tracker not found at {TRACKER_PATH}.\n"
            "Run send_survey_emails.py first to build the tracker, then retry.\n"
        )
        # Still continue — we can still send reminders if the tracker exists
        # but happens to have been created outside the data/ folder.
        # Create an empty tracker so the script doesn't crash.
        tracker_df = pd.DataFrame(columns=config.get("tracker_columns", []))
    else:
        tracker_df = load_tracker()

    # Ensure all required columns exist (may be absent in older tracker files)
    for col in ["survey_completed_at", "follow_up_1_sent_at", "follow_up_2_sent_at"]:
        if col not in tracker_df.columns:
            tracker_df[col] = ""

    # ── Process response ZIPs ─────────────────────────────────────────────────
    if not zips:
        print("No .zip files found in the ingest/ folder.")
        print("  • If you have a response survey export to process, drop its ZIP there.")
        print("  • The reminder check will still run based on the existing tracker.\n")
    else:
        print(f"Found {len(zips)} ZIP(s): {', '.join(z.name for z in zips)}\n")
        tracker_df, flags = process_response_zips(zips, tracker_df)
        if flags:
            print(f"Flagged issues: {len(flags)}")
            for f in flags:
                print(f"  • [{f['zip']}] {f['issue']}  (response_id={f['response_id']})")
            print()

    now_iso = datetime.now().isoformat(timespec="seconds")

    # Count completions for the summary (gift-card handling moved to
    # payment_management/manage_payments.py).
    completed_count = int(
        tracker_df["survey_completed_at"].str.strip().ne("").sum()
    ) if "survey_completed_at" in tracker_df.columns else 0

    # ── Reminder check ────────────────────────────────────────────────────────
    print("═" * 56)
    print("  REMINDER CHECK")
    print("═" * 56)
    tracker_df, send_log_rows, reminder_log_rows, reminder_counts = \
        check_and_send_reminders(tracker_df, draft_only, dry_run, now_iso)

    if not any(reminder_counts.values()):
        print("  No reminders due at this time.\n")
    else:
        print()

    # ── Persist ───────────────────────────────────────────────────────────────
    save_tracker(tracker_df)
    append_send_log(send_log_rows)
    append_reminder_log(reminder_log_rows)

    # ── Summary ───────────────────────────────────────────────────────────────
    print("═" * 56)
    print(f"  SUMMARY  ({mode_label})")
    print("═" * 56)
    print(f"  Completions in tracker : {completed_count}")
    print(f"  (Run payment_management to manage gift cards for these completions.)")
    for status, n in sorted(reminder_counts.items()):
        print(f"  {status:<30}: {n}")
    print(f"  Tracker saved to       : {TRACKER_PATH}")
    if reminder_log_rows:
        print(f"  Reminder log           : {REMINDER_LOG_PATH}  "
              f"(+{len(reminder_log_rows)} this run)")
    if draft_only:
        where = (f"the '{SHARED_MAILBOX}' Drafts folder"
                 if SHARED_MAILBOX else "your Outlook Drafts")
        print(f"  Reminder emails staged in {where}.")
    if dry_run:
        print("  Dry run — no emails sent, tracker updated only.")
    print()


if __name__ == "__main__":
    main()
