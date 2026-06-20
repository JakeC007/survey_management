"""
send_survey_emails.py
─────────────────────
DROP-A-ZIP PIPELINE for the Teen AI Survey invitations.

How you use it (the whole flow)
───────────────────────────────
  1. Download the Qualtrics screener export ZIP.
  2. Drop it into the  ingest  folder (survey_management/ingest/;
     created automatically the first time you run).
  3. Double-run the script:

         python send_survey_emails.py

     That's it. No file paths, no flags. It will:
       • read the MOST RECENT .zip in ingest (older ZIPs are left in place)
       • record EVERY response in the master tracker
       • email only people it has NOT emailed before (dedup via the tracker)
       • stamp each first send with a date/time in the tracker
       • print + save run stats
       • keep every ZIP (the inbox is an append-only archive)

Safety flags (optional)
───────────────────────
  --draft     Stage emails as Outlook Drafts for review instead of sending.
              Recommended for your FIRST real batch.
  --dry-run   Parse + build the tracker only. No Outlook, no ZIP deletion.
              Use this to sanity-check a new export.

Outputs (all inside data/ subfolder)
─────────────────────────────────────
  data/participant_tracker_auto.xlsx  — ONE master sheet, upserted every run.
                                        One row per Response ID, every response
                                        ever ingested, with status, reason, grade,
                                        and emailed_at (first-send timestamp).
                                        Plus a "Run Stats" sheet.
                                        Response-tracking columns (survey_completed_at,
                                        follow_up_1_sent_at, follow_up_2_sent_at) are
                                        written by manage_responses.py and preserved here.
  data/send_log.csv                   — ONE append-only file. One row per email
                                        actually sent/drafted. Backs up the dedup.

Requirements
────────────
  pip install pywin32 pandas openpyxl
  Outlook must be open and logged in as your UChicago account
  (not needed for --dry-run).
"""

# Bot/fraud/duplicate screening lives in flag_suspicious_entries.py and is
# imported lazily inside process_dataframe (see note there) to avoid a circular
# import, since that module imports the ZIP/column helpers from this one.
import argparse
import sys
import zipfile
import io
import os
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml


SCRIPT_DIR  = Path(__file__).parent          # …/survey_management/consent_management/
REPO_ROOT   = SCRIPT_DIR.parent             # …/survey_management/

CONFIG_PATH = REPO_ROOT / "config.yaml"

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

# ── GENERAL SETTINGS ───────────────────────────────────────────────────────────

SURVEY_LINK = config["survey_link"]

SENDER_NAME = config["sender"]["name"]
SENDER_TITLE = config["sender"]["title"]
CONTACT_EMAIL = config["sender"]["contact_email"]

SHARED_MAILBOX = config["shared_mailbox"]

# ── PATHS ──────────────────────────────────────────────────────────────────────
# All paths resolve from REPO_ROOT (survey_management/) so this script shares
# the ingest folder and bookkeeping files with manage_responses.py.

INBOX_DIR    = REPO_ROOT / config["paths"]["inbox_dir"]
TRACKER_PATH = REPO_ROOT / config["paths"]["tracker_file"]
SEND_LOG_PATH = REPO_ROOT / config["paths"]["send_log_file"]

# ── QUALTRICS FIELD NAMES ──────────────────────────────────────────────────────

fields = config["qualtrics_fields"]

F_RESPONSE_ID = fields["response_id"]
F_RECORDED_DATE = fields["recorded_date"]
F_FINISHED = fields["finished"]
F_PARENT_NAME = fields["parent_name"]
F_CHILD_NAME = fields["child_name"]
F_PARENT_CONSENT = fields["parent_consent"]
F_PARENT_SIG = fields["parent_signature"]
F_CHILD_ASSENT = fields["child_assent"]
F_CHILD_SIG = fields["child_signature"]
F_DELIVERY_EMAIL = fields["delivery_email"]

# ── ACCEPTABLE RESPONSES ───────────────────────────────────────────────────────

CONSENT_OK = set(config["consent_ok"])
ASSENT_OK = set(config["assent_ok"])

# ── TRACKER COLUMNS ────────────────────────────────────────────────────────────

TRACKER_COLS = config["tracker_columns"]


# ── HELPERS ───────────────────────────────────────────────────────────────────

def load_csv_from_zip(zip_path: str) -> pd.DataFrame:
    """Extract the first CSV from a Qualtrics export ZIP and return the data rows
    with the human-readable LABEL row as the column header.

    A standard Qualtrics export has three header rows:
      row 0 = short variable names  (StartDate, ResponseId, Q3, Q4, ...)
      row 1 = human-readable labels (Start Date, Response ID, "Parent/Guardian...")
      row 2 = ImportId JSON         ({"ImportId":"..."} ...)
    then the actual responses. This script matches on the labels, so we use
    row 1 as the header and drop the ImportId row. We also handle the simpler
    two-row export (labels in row 0) by locating whichever row holds the labels.
    """
    with zipfile.ZipFile(zip_path) as z:
        csv_names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not csv_names:
            raise ValueError("No CSV found inside the ZIP.")
        raw = z.read(csv_names[0])

    grid = pd.read_csv(io.BytesIO(raw), dtype=str, header=None, encoding="utf-8-sig")

    # Find the label row: the first of the top rows containing "Response ID".
    label_row = 0
    for i in range(min(3, len(grid))):
        if grid.iloc[i].map(lambda v: _norm(v) == "response id").any():
            label_row = i
            break

    grid.columns = grid.iloc[label_row]
    data = grid.iloc[label_row + 1:]
    # Drop the ImportId metadata row if it directly follows the labels.
    if len(data) and data.iloc[0].map(lambda v: str(v).startswith('{"ImportId"')).any():
        data = data.iloc[1:]
    return data.reset_index(drop=True).fillna("")


def _norm(s) -> str:
    """Normalise a label for tolerant matching: lowercase, unify smart quotes
    and dashes, collapse all whitespace (incl. embedded newlines) to one space."""
    s = "" if s is None else str(s)
    for a, b in (("’", "'"), ("‘", "'"), ("“", '"'), ("”", '"'),
                 ("–", "-"), ("—", "-")):
        s = s.replace(a, b)
    return " ".join(s.lower().split())


# ── Email-domain hygiene (shared by the mailer and the consent screener) ──────
# Common misspellings of big providers' domains. These are CORRECTED silently so
# a real family that typo'd their address still gets the survey, instead of being
# bounced or flagged. Gmail is the bulk of real-world slips.
EMAIL_TYPO_FIXES = {
    "gmial.com": "gmail.com", "gmil.com": "gmail.com", "gmai.com": "gmail.com",
    "gmal.com": "gmail.com", "gamil.com": "gmail.com", "gnail.com": "gmail.com",
    "gmaill.com": "gmail.com", "gmail.co": "gmail.com", "gmail.cm": "gmail.com",
    "gmail.con": "gmail.com", "gmail.comm": "gmail.com", "gmail.om": "gmail.com",
    "gmailo.com": "gmail.com", "gmaul.com": "gmail.com", "gmali.com": "gmail.com",
    "yaho.com": "yahoo.com", "yahooo.com": "yahoo.com", "yhoo.com": "yahoo.com",
    "hotmial.com": "hotmail.com", "hotmal.com": "hotmail.com", "hotmai.com": "hotmail.com",
    "outlok.com": "outlook.com", "outloo.com": "outlook.com",
    "iclod.com": "icloud.com", "icloud.co": "icloud.com",
}


def email_domain(email) -> str:
    """Lowercased domain part of an email, or '' if there isn't a clean one."""
    e = (email or "").strip().lower()
    return e.rsplit("@", 1)[1] if e.count("@") == 1 else ""


def correct_email_typo(email):
    """Return (corrected_email, was_corrected). Fixes a known typo'd provider
    domain (mostly gmail) so a slip like 'gmil.com' still reaches a real family.
    Leaves anything not in the map untouched."""
    e = (email or "").strip()
    if e.count("@") != 1:
        return e, False
    local, dom = e.rsplit("@", 1)
    fixed = EMAIL_TYPO_FIXES.get(dom.lower())
    if fixed and fixed != dom.lower():
        return f"{local}@{fixed}", True
    return e, False


# Each field is matched to its column by a distinctive normalised substring,
# so smart quotes, em dashes, and line breaks in Qualtrics labels can't break it.
FIELD_MATCHERS = {
    "response_id":    "response id",
    "recorded_date":  "recorded date",
    "finished":       "finished",
    "parent_name":    "parent/guardian full name",
    "child_name":     "child participant's full name",
    "parent_consent": "i confirm that i have read",
    "parent_sig":     "typed signature - parent",
    "child_assent":   "i understand what the researchers asked",
    "child_sig":      "typed signature - child",
    "delivery_email": "what email address should we send",
    "grade":          "what grade was the child in",
}

# Fields that may be absent from an export (e.g. questions added partway through
# data collection). resolve_columns won't raise if these are missing; the row
# loop records a blank value instead of failing the whole batch.
OPTIONAL_FIELDS = {"grade"}


# A consent/screener export is told apart from the teen response export by a
# question stem that only exists on the consent survey. Cheap heuristic: peek at
# the first few header rows and look for any of these. The inbox holds both
# export types, so each script picks the newest file of the type it wants.
CONSENT_ONLY_MARKERS = (
    "i confirm that i have read",
    "i understand what the researchers asked",
    "parent/guardian full name",
)


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


def looks_like_consent_export(path) -> bool:
    """True if this ZIP's headers match the consent/screener survey."""
    try:
        text = _peek_header_text(path)
    except Exception:
        return False
    return any(m in text for m in CONSENT_ONLY_MARKERS)


def pick_consent_zip(zips):
    """Most recently MODIFIED ZIP whose headers match the consent schema."""
    for zp in sorted(zips, key=lambda p: p.stat().st_mtime, reverse=True):
        if looks_like_consent_export(zp):
            return zp
    return None


def resolve_columns(df: pd.DataFrame) -> dict:
    """Map each logical field to the actual column name in this export.
    Raises if a required column can't be found, so problems surface loudly
    rather than silently skipping everyone."""
    norm_cols = [(_norm(c), c) for c in df.columns]
    resolved, missing = {}, []
    for field, needle in FIELD_MATCHERS.items():
        match = next((orig for n, orig in norm_cols if needle in n), None)
        if match is None:
            if field not in OPTIONAL_FIELDS:
                missing.append(f"{field} (looked for '{needle}')")
        else:
            resolved[field] = match
    if missing:
        raise KeyError("Could not find these columns in the export: "
                       + "; ".join(missing))
    return resolved


def get_child_first_name(full_name: str) -> str:
    parts = full_name.strip().split()
    return parts[0] if parts else full_name


def load_tracker() -> dict:
    """Load existing master tracker into a dict keyed by response_id."""
    if not TRACKER_PATH.exists():
        return {}
    df = pd.read_excel(TRACKER_PATH, sheet_name="Participants", dtype=str).fillna("")
    return {r["response_id"]: dict(r) for _, r in df.iterrows() if r.get("response_id")}


def emailed_ids_from_tracker(tracker: dict) -> set:
    """Response IDs already emailed = those with a non-blank emailed_at in the tracker.
    Cross-checked against send_log.csv as a backup source of truth."""
    ids = {rid for rid, r in tracker.items() if str(r.get("emailed_at", "")).strip()}
    if SEND_LOG_PATH.exists():
        df = pd.read_csv(SEND_LOG_PATH, dtype=str).fillna("")
        ids |= set(df.loc[df["status"].isin(["SENT", "DRAFTED"]), "response_id"])
    return ids


def emailed_at_from_send_log() -> dict:
    """Map response_id -> emailed_at from send_log.csv (first SENT/DRAFTED row).
    Lets the tracker recover the send timestamp for already-emailed people even
    if the tracker file itself was lost or never written on a previous run."""
    if not SEND_LOG_PATH.exists():
        return {}
    out = {}
    try:
        df = pd.read_csv(SEND_LOG_PATH, dtype=str).fillna("")
    except Exception:
        return {}
    for _, r in df.iterrows():
        rid = str(r.get("response_id", "")).strip()
        if rid and r.get("status") in ("SENT", "DRAFTED"):
            out.setdefault(rid, str(r.get("emailed_at", "")).strip())
    return out


def _recovery_path(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return path.with_name(f"{path.stem}.RECOVERY-{stamp}{path.suffix}")


def append_send_log(rows: list):
    """Append rows to the single send_log.csv, writing a header only if new.

    If the file can't be written (e.g. it's open in Excel), the records are NOT
    lost: they go to a timestamped RECOVERY file and the run is warned loudly,
    so the fact that those emails went out is never silently dropped.
    """
    if not rows:
        return
    SEND_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)   # ensure data/ exists
    df = pd.DataFrame(rows, columns=["response_id", "child_name", "delivery_email",
                                     "status", "mode", "emailed_at"])
    try:
        df.to_csv(SEND_LOG_PATH, mode="a", header=not SEND_LOG_PATH.exists(),
                  index=False, encoding="utf-8")
    except Exception as e:
        fb = _recovery_path(SEND_LOG_PATH)
        df.to_csv(fb, index=False, encoding="utf-8")
        print(f"\n  !! Could NOT write {SEND_LOG_PATH.name} ({e}).")
        print(f"     This run's {len(rows)} send record(s) were saved to {fb.name} instead.")
        print(f"     Close {SEND_LOG_PATH.name} in Excel, then append {fb.name} into it")
        print(f"     BEFORE the next run — otherwise those people may be emailed again.\n")


def write_tracker(tracker: dict, stats_rows: list):
    """Write the master tracker (Participants sheet) + a Run Stats sheet.

    Like append_send_log, on a write failure (file open in Excel) it falls back
    to a timestamped RECOVERY copy instead of crashing, so the run's bookkeeping
    is never lost after emails have already gone out.
    """
    TRACKER_PATH.parent.mkdir(parents=True, exist_ok=True)   # ensure data/ exists
    df = pd.DataFrame(list(tracker.values()))
    for c in TRACKER_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[TRACKER_COLS].sort_values("recorded_date")
    stats_df = pd.DataFrame(stats_rows, columns=["metric", "count"])

    def _write(path):
        with pd.ExcelWriter(path, engine="openpyxl") as xl:
            df.to_excel(xl, sheet_name="Participants", index=False)
            stats_df.to_excel(xl, sheet_name="Run Stats", index=False)

    try:
        _write(TRACKER_PATH)
    except Exception as e:
        fb = _recovery_path(TRACKER_PATH)
        _write(fb)
        print(f"\n  !! Could NOT write {TRACKER_PATH.name} ({e}).")
        print(f"     The tracker for this run was saved to {fb.name} instead.")
        print(f"     Close {TRACKER_PATH.name} in Excel and replace it with {fb.name}.\n")

def generate_user_url(rid):
    unique_url = f"{SURVEY_LINK}?cid={rid}"
    return unique_url

def build_email_body(child_first, sender_name, sender_title, contact, unique_url) -> str:
    return f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
<p>Hi {child_first},</p>

<p>We are a team of researchers from the University of Chicago studying how students like you
feel about artificial intelligence (AI) and privacy in schools. As AI tools become more common
in classrooms — from tutoring software to automated grading — we want to hear directly from
students about their experiences and concerns.</p>

<p>Your parent or guardian has given permission for you to participate. We're inviting you to
complete a short online survey that should take no more than 10–15 minutes.</p>

<p style="margin:20px 0;">
  <a href="{unique_url}" style="background:#800000;color:#fff;padding:10px 20px;
  text-decoration:none;border-radius:4px;font-weight:bold;">Take the Survey →</a>
</p>

<p>Or copy this link into your browser:<br>
<a href="{unique_url}">{unique_url}</a></p>

<p><strong>Why participate?</strong><br>
Your voice matters — student perspectives are rarely included in conversations about school
technology policy. You'll also receive a <strong>$10 Amazon gift card</strong> as a thank-you
for completing the survey.</p>

<p>Your responses will be kept completely confidential and used only for research purposes.
At the start of the survey you'll find a brief assent form confirming you understand what
the study involves.</p>

<p>Thank you — we really do want to hear from you.</p>

<p>Best regards,<br>
<strong>{sender_name}</strong><br>
{sender_title}<br>
Investigations of Educational Technology to Safeguard Children's Privacy<br>
University of Chicago<br>
<a href="mailto:{contact}">{contact}</a></p>
</body></html>"""


_OUTLOOK = {}   # caches the Outlook app + resolved shared Drafts folder for the run


def _get_outlook():
    import win32com.client as win32
    if "app" not in _OUTLOOK:
        _OUTLOOK["app"] = win32.Dispatch("outlook.application")
    return _OUTLOOK["app"]


def _shared_drafts_folder():
    """The shared mailbox's Drafts folder, or None if no shared mailbox is set.
    Needs Full Access to the shared mailbox (not just Send As)."""
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
    try:
        folder = ns.GetSharedDefaultFolder(recip, OL_FOLDER_DRAFTS)
    except Exception as e:
        raise RuntimeError(
            f"Could not open the Drafts folder of '{SHARED_MAILBOX}'. This usually "
            "means you have Send-As but not Full Access to that mailbox. Details: " + str(e))
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


def preflight(draft_only, dry_run, record_only=False):
    """Fail fast BEFORE touching any recipient.

    Two checks, in this order:
      1. The bookkeeping files (send_log.csv, tracker) are writable. This is the
         critical one: we must never send a batch of emails and then be unable
         to record who got them. Checked in every mode, since all modes write
         the tracker at the end.
      2. Outlook is reachable (skipped for dry-run and record-only, which send
         nothing).
    """
    _assert_writable(SEND_LOG_PATH, "send_log.csv")
    _assert_writable(TRACKER_PATH, "the participant tracker")
    if dry_run or record_only:
        return
    try:
        _get_outlook()
    except Exception as e:
        raise RuntimeError(
            "Could not connect to Outlook. Make sure classic Outlook (NOT New "
            "Outlook) is open and signed in. Details: " + str(e))
    # If we'll store drafts in the shared mailbox, confirm we can reach the folder now.
    if SHARED_MAILBOX and draft_only:
        _shared_drafts_folder()


def send_via_outlook(to_email, subject, html_body, draft_only):
    mail = _get_outlook().CreateItem(0)  # olMailItem
    mail.To = to_email
    mail.Subject = subject
    mail.HTMLBody = html_body
    # Route the From to a shared mailbox if configured. With "Send As"
    # permission recipients see the mail as from the shared box; with
    # "Send on Behalf" they see "you on behalf of the shared box".
    if SHARED_MAILBOX:
        mail.SentOnBehalfOfName = SHARED_MAILBOX
    if draft_only:
        mail.Save()  # lands in your default Drafts first
        folder = _shared_drafts_folder()
        if folder is not None:
            mail.Move(folder)  # relocate into the shared mailbox's Drafts
    else:
        mail.Send()


# ── PER-ZIP PROCESSING ────────────────────────────────────────────────────────

def reason_category(r: str) -> str:
    if r.startswith("consent"):       return "consent not given"
    if "parent signature" in r:       return "missing parent signature"
    if r.startswith("assent"):        return "assent not given"
    if "child signature" in r:        return "missing child signature"
    if "email" in r:                  return "bad/blank email"
    return "other"


def process_dataframe(df, tracker, already_emailed, mode_label, draft_only, dry_run,
                      now_iso, record_only=False, emailed_at_map=None):
    """Process one export's rows: upsert tracker, send/draft eligible, collect stats.
    Mutates `tracker` and `already_emailed` in place. Returns per-run artefacts.

    record_only: don't touch Outlook; mark eligible rows as SENT and record them.
        Use this to rebuild send_log.csv / tracker after emails went out but the
        bookkeeping write failed (see --record-only).
    emailed_at_map: response_id -> emailed_at from send_log.csv, used to restore
        the send timestamp onto ALREADY_SENT rows when the tracker was lost."""
    emailed_at_map = emailed_at_map or {}
    send_log_rows = []
    status_counts = Counter()
    reason_counts = Counter()
    emailed_to = []

    cols = resolve_columns(df)   # logical field -> actual column name in this export

    # Screen the whole batch for bot/fraud/duplicate signals up front. Imported
    # lazily because flag_suspicious_entries imports helpers from this module
    # (a top-level import here would be circular). flagged maps response_id ->
    # record; suspicious_flags is the subset to skip and label SUSPICIOUS.
    from flag_suspicious_entries import flag_entries, write_blacklist, FRAUD_BLACKLIST_PATH
    flagged = {r["response_id"]: r for r in flag_entries(df)}
    suspicious_flags = {rid for rid, r in flagged.items() if r["suspicious"]}
    if suspicious_flags:
        print(f"  Suspicious-entry screen: {len(suspicious_flags)} flagged "
              f"(will be recorded SUSPICIOUS and NOT emailed).")
        # Persist these to the shared fraud blacklist so payment_management can
        # refuse a gift card to anyone whose cid or IP is on the list, even if
        # the survey link somehow reaches them. Append-only; manual rows kept.
        # Runs in every mode (including --dry-run) so a dry pass still populates it.
        try:
            added = write_blacklist([flagged[r] for r in suspicious_flags])
            print(f"  Fraud blacklist updated: +{added} new "
                  f"({FRAUD_BLACKLIST_PATH.name}).")
        except Exception as e:
            print(f"  WARNING: could not update fraud blacklist: {e}")

    for _, row in df.iterrows():
        rid = str(row.get(cols["response_id"], "")).strip()
        if not rid:
            continue

        parent_name    = str(row.get(cols["parent_name"], "")).strip()
        child_name     = str(row.get(cols["child_name"], "")).strip()
        parent_consent = str(row.get(cols["parent_consent"], "")).strip().lower()
        parent_sig     = str(row.get(cols["parent_sig"], "")).strip()
        child_assent   = str(row.get(cols["child_assent"], "")).strip().lower()
        child_sig      = str(row.get(cols["child_sig"], "")).strip()
        delivery_email = str(row.get(cols["delivery_email"], "")).strip()
        delivery_email, _email_fixed = correct_email_typo(delivery_email)
        if _email_fixed:
            print(f"  CORRECTED email typo → {delivery_email}")
        # Grade is a later-added, optional question. The column may be absent on
        # older exports, and the value may be blank on responses recorded before
        # the question went live — both are fine, we just note it as missing and
        # never block the invitation on it.
        grade          = str(row.get(cols.get("grade", ""), "")).strip()
        recorded_date  = str(row.get(cols["recorded_date"], "")).strip()
        finished       = str(row.get(cols["finished"], "")).strip().upper() in ("TRUE", "1")

        status = ""
        reason = ""
        emailed_at = ""

        if not finished:
            status = "INCOMPLETE"
            reason = "Qualtrics Finished != True"
        elif rid in already_emailed:
            status = "ALREADY_SENT"
            reason = "already in send_log.csv"
            emailed_at = emailed_at_map.get(rid, "")   # restore timestamp if tracker was lost
        elif rid in suspicious_flags:
            status = "SUSPICIOUS"
            reason = flagged[rid]["reasons"]
            reason_counts["suspicious entry"] += 1
            print(f"  FLAG  {child_name or rid}: {reason}")
        else:
            reasons = []
            if parent_consent not in CONSENT_OK:
                reasons.append(f"consent='{parent_consent}'")
            if not parent_sig:
                reasons.append("missing parent signature")
            if child_assent not in ASSENT_OK:
                reasons.append(f"assent='{child_assent}'")
            if not child_sig:
                reasons.append("missing child signature")
            if not delivery_email or "@" not in delivery_email:
                reasons.append(f"bad email='{delivery_email}'")
            # NOTE: Age is no longer gated here. Age eligibility is verified
            # upstream in the consent form, and DOB is no longer collected.

            if reasons:
                status = "INELIGIBLE"
                reason = "; ".join(reasons)
                for r in reasons:
                    reason_counts[reason_category(r)] += 1
                print(f"  SKIP  {child_name or rid}: {reason}")
            else:
                child_first = get_child_first_name(child_name)
                subject = "You're Invited: Share Your Thoughts on AI in School & Earn $10"
                unique_url = generate_user_url(rid)
                body = build_email_body(child_first, SENDER_NAME, SENDER_TITLE, CONTACT_EMAIL, unique_url)
                if dry_run:
                    status = "ELIGIBLE (dry-run)"
                    print(f"  WOULD EMAIL  {child_first} → {delivery_email}")
                elif record_only:
                    # Emails already went out on a previous run; just record them
                    # so dedup and the tracker are correct. No Outlook contact.
                    status = "SENT"
                    emailed_at = datetime.now().isoformat(timespec="seconds")
                    emailed_to.append((child_name, delivery_email))
                    already_emailed.add(rid)
                    send_log_rows.append({
                        "response_id": rid, "child_name": child_name,
                        "delivery_email": delivery_email, "status": status,
                        "mode": mode_label, "emailed_at": emailed_at,
                    })
                    print(f"  RECORDED (already sent)  {child_first} → {delivery_email}")
                else:
                    try:
                        send_via_outlook(delivery_email, subject, body, draft_only)
                        status = "DRAFTED" if draft_only else "SENT"
                        emailed_at = datetime.now().isoformat(timespec="seconds")
                        emailed_to.append((child_name, delivery_email))
                        already_emailed.add(rid)   # dedup within this run too
                        send_log_rows.append({
                            "response_id": rid, "child_name": child_name,
                            "delivery_email": delivery_email, "status": status,
                            "mode": mode_label, "emailed_at": emailed_at,
                        })
                        print(f"  {status}  {child_first} → {delivery_email}")
                    except Exception as e:
                        status = "ERROR"
                        reason = str(e)
                        print(f"  ERROR  {child_first} → {delivery_email}: {e}")

        status_counts[status] += 1

        # Upsert into master tracker
        existing = tracker.get(rid, {})
        tracker[rid] = {
            "response_id":   rid,
            "recorded_date": recorded_date,
            "parent_name":   parent_name,
            "child_name":    child_name,
            "delivery_email": delivery_email,
            "grade":         grade or "(missing)",
            "status":        status,
            "reason":        reason,
            "first_seen":    existing.get("first_seen") or now_iso,
            "last_updated":  now_iso,
            "emailed_at":    emailed_at or existing.get("emailed_at", ""),
            "last_mode":     mode_label,
            # ── Preserve response-management columns ──────────────────────────
            # These columns are written by manage_responses.py and must NOT be
            # overwritten by the invitation script on re-runs.
            "survey_completed_at": existing.get("survey_completed_at", ""),
            "follow_up_1_sent_at": existing.get("follow_up_1_sent_at", ""),
            "follow_up_2_sent_at": existing.get("follow_up_2_sent_at", ""),
        }

    return send_log_rows, status_counts, reason_counts, emailed_to, rid


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Drop a Qualtrics screener ZIP in the ingest folder, then run this.")
    parser.add_argument("--draft", action="store_true",
                        help="Stage emails as Outlook Drafts instead of sending.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse + build tracker only. No Outlook, no ZIP deletion.")
    parser.add_argument("--inbox", default=str(INBOX_DIR),
                        help=f"Folder to scan for ZIPs (default: {INBOX_DIR.name}).")
    parser.add_argument("--record-only", action="store_true",
                        help="Recovery mode: do NOT contact Outlook. Mark every "
                             "eligible row as SENT and (re)write send_log.csv + the "
                             "tracker. Use this when emails already went out but the "
                             "bookkeeping write failed, so dedup is restored.")
    args = parser.parse_args()

    dry_run = args.dry_run
    draft_only = args.draft
    record_only = args.record_only
    mode_label = ("DRY-RUN" if dry_run else "RECORD-ONLY" if record_only
                  else "DRAFT" if draft_only else "SEND")

    inbox = Path(args.inbox)
    inbox.mkdir(parents=True, exist_ok=True)
    all_zips = sorted(inbox.glob("*.zip"), key=lambda p: p.stat().st_mtime)

    print(f"\nMode: {mode_label}")
    print(f"Inbox: {inbox}")
    if not all_zips:
        print(f"\nNo .zip files found in {inbox.name}.")
        print("Drop your Qualtrics screener export there and run again.\n")
        return

    # The inbox holds both consent and response exports. Pick the most recently
    # MODIFIED ZIP whose headers match the consent/screener schema; quietly skip
    # response-survey ZIPs. Dedup against the tracker means re-reading an
    # already-processed export never re-emails anyone.
    chosen = pick_consent_zip(all_zips)
    if chosen is None:
        print(f"\nNo consent/screener export found among {len(all_zips)} ZIP(s) "
              f"in {inbox.name}.")
        print("(Checked headers; none matched the consent survey schema.)")
        print("Drop your Qualtrics screener export there and run again.\n")
        return
    zips = [chosen]
    others = [z.name for z in all_zips if z != chosen]
    print(f"Using consent export: {chosen.name}")
    if others:
        print(f"  (ignoring {len(others)} other ZIP(s): {', '.join(others)})")
    print()

    # Fail fast if Outlook (or the shared Drafts folder) isn't usable, before
    # we process anyone. Avoids marking a whole batch ERROR against a dead Outlook.
    try:
        preflight(draft_only, dry_run, record_only)
    except Exception as e:
        print(f"PREFLIGHT FAILED: {e}\n")
        return

    now_iso = datetime.now().isoformat(timespec="seconds")
    tracker = load_tracker()
    already_emailed = emailed_ids_from_tracker(tracker)
    emailed_at_map = emailed_at_from_send_log()   # restore send timestamps if tracker was lost
    if already_emailed:
        print(f"Dedup: {len(already_emailed)} people already emailed — they'll be skipped.\n")

    all_send_log = []
    total_status = Counter()
    total_reason = Counter()
    all_emailed_to = []
    processed_zips = []
    errored_zips = []
    failed_zips = []
    total_rows = 0

    for zp in zips:
        print(f"── Ingesting {zp.name} ──")
        try:
            df = load_csv_from_zip(str(zp))
            slr, sc, rc, et, rid = process_dataframe(
                df, tracker, already_emailed, mode_label, draft_only, dry_run, now_iso,
                record_only=record_only, emailed_at_map=emailed_at_map)
        except Exception as e:
            print(f"  COULD NOT PROCESS {zp.name}: {e}  (left in inbox)\n")
            failed_zips.append(zp)
            continue

        total_rows += len(df)
        all_send_log += slr
        total_status += sc
        total_reason += rc
        all_emailed_to += et
        processed_zips.append(zp)
        if sc.get("ERROR", 0):   # a send failed in this batch — keep the ZIP for retry
            errored_zips.append(zp)
        print()

    # ── Persist BEFORE deleting any ZIP ───────────────────────────────────────
    append_send_log(all_send_log)
    stats_rows = [("run_timestamp", now_iso), ("mode", mode_label),
                  ("zips_processed", len(processed_zips)),
                  ("rows_in_exports", total_rows)]
    for k in ["SENT", "DRAFTED", "ELIGIBLE (dry-run)", "SUSPICIOUS", "INELIGIBLE",
              "INCOMPLETE", "ALREADY_SENT", "ERROR"]:
        if total_status.get(k):
            stats_rows.append((f"status:{k}", total_status[k]))
    for cat, n in total_reason.most_common():
        stats_rows.append((f"ineligible_reason:{cat}", n))
    write_tracker(tracker, stats_rows)

    # ── ZIPs are kept, never deleted ──────────────────────────────────────────
    # The inbox is treated as an append-only archive. Each run reads only the
    # most recent ZIP; older exports are left in place for your records. Dedup
    # against the tracker (emailed_ids) guarantees no one is emailed twice even
    # if the same export is read again.

    # ── Console stats ─────────────────────────────────────────────────────────
    print("═" * 52)
    print(f"  RUN STATS  ({mode_label})")
    print("═" * 52)
    emailed_n = total_status.get("SENT", 0) + total_status.get("DRAFTED", 0)
    verb = {"SEND": "Sent", "DRAFT": "Drafted",
            "RECORD-ONLY": "Recorded (already sent)"}.get(mode_label, "Would email")
    would = total_status.get("ELIGIBLE (dry-run)", 0)
    print(f"  {verb:<22}: {emailed_n or would}")
    for name, email in all_emailed_to:
        print(f"      → {name or '(no name)'}  <{email}>")
    print(f"  Suspicious (skipped)  : {total_status.get('SUSPICIOUS', 0)}")
    print(f"  Ineligible (skipped)  : {total_status.get('INELIGIBLE', 0)}")
    for cat, n in total_reason.most_common():
        print(f"      • {cat}: {n}")
    print(f"  Incomplete responses  : {total_status.get('INCOMPLETE', 0)}")
    print(f"  Already emailed (skip): {total_status.get('ALREADY_SENT', 0)}")
    print(f"  Errors                : {total_status.get('ERROR', 0)}")
    print("=" * 52)
    print(f"  Tracker : {TRACKER_PATH.name}  ({len(tracker)} participants total)")
    print(f"  Send log: {SEND_LOG_PATH.name}  (+{len(all_send_log)} this run)")
    print(f"  Inbox    : ZIPs are kept (newest is read each run, none deleted).")
    if failed_zips:
        print(f"  Unreadable this run: {', '.join(z.name for z in failed_zips)}")
    if errored_zips:
        print(f"  Note: send errors occurred — re-run to retry "
              f"(dedup skips anyone already emailed).")
    if draft_only:
        where = f"the '{SHARED_MAILBOX}' Drafts folder" if SHARED_MAILBOX else "your Outlook Drafts"
        print(f"  Emails are in {where}. Review and send from there.")
    if dry_run:
        print("  Dry run: no emails sent.")
    if record_only:
        print("  Record-only: no emails sent. send_log.csv + tracker rebuilt so "
              "dedup is restored.")
    print()


if __name__ == "__main__":
    main()
