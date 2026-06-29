"""
manage_payments.py
──────────────────
Payment / gift-card management pipeline for the Teen AI Survey.

This is the third leg of the survey-management system, alongside
consent_management/ (invites) and response_management/ (completions +
reminders).  It owns everything about WHO GETS PAID.  The gift-card list that
used to print from manage_responses.py now lives here.

WHAT IT DOES
────────────
  1. Reads the shared master tracker (data/participant_tracker_auto.xlsx) — the
     SAME data source the other two folders use.  A participant is payable when
     they have BOTH consent AND a completed survey.  Both facts already live in
     the tracker: each row's response_id IS the consent Response ID (= cid), and
     survey_completed_at is stamped by manage_responses.py only when a completed
     response's cid matches that row.  So "consent AND complete, merged on cid"
     reduces to: a tracker row whose survey_completed_at is non-empty.

  2. Runs the shared quality_filter.py against the RESPONSE-survey export (the
     Qualtrics teen-survey ZIP/CSV in ingest/).  quality_filter keys each
     respondent by cid, so its flags merge straight onto the tracker rows.

  3. Upserts a master payment ledger (data/payment_tracker.xlsx).  One row per
     completed participant, carrying the quality flags AND a hand-editable
     `paid` column.  Re-runs NEVER overwrite your manual paid/notes edits — they
     are merged back in by cid.  This is the sheet you work from when you buy
     and send Amazon cards by hand.

  4. Exports the unpaid report (data/payment_report_unpaid.xlsx) — everyone who
     completed but is not yet marked paid, split into three sheets:
        • "Pay"   — no exclusion recommended by the quality filter.
        • "Hold"  — exclude_recommended by the quality filter (review first).
        • "Fraud" — cid or IP is on the fraud blacklist. NEVER pay these.
     Each row carries first name + the delivery email the survey was sent to,
     pulled from the tracker (which copied them from the consent data).

  ┌─ FRAUD BUCKET vs HOLD BUCKET ────────────────────────────────────────────┐
  │ These are different things and are decided independently:                 │
  │   • HOLD is a DATA-QUALITY judgement on a real participant's answers       │
  │     (speeding, failed attention checks, straight-lining). Held teens may   │
  │     still be paid after review; exclusion is an analysis decision.         │
  │   • FRAUD is an IDENTITY judgement made upstream at consent time by        │
  │     consent_management's suspicious-entry screen (bot/fraud-farm/duplicate │
  │     signals). Those response IDs and their IPs are written to              │
  │     data/fraud_blacklist.csv. A completer is routed to Fraud — and never   │
  │     paid — if its cid OR its survey IP is on that list. This catches a     │
  │     scammer who completes the survey after somehow obtaining a link.       │
  │ Fraud takes priority: a fraud-matched row never appears in Pay or Hold.    │
  └───────────────────────────────────────────────────────────────────────────┘

  ┌─ NOTE ON THE HOLD BUCKET ────────────────────────────────────────────────┐
  │ QUALITY_CONTROL.md states the default protocol is to PAY every completer  │
  │ and treat exclusion as an analysis-only decision.  Splitting unpaid teens │
  │ into a Hold bucket conditions payment on quality flags, which deviates    │
  │ from that protocol.  This script does it because the project owner        │
  │ confirmed the deviation was cleared with the IRB.  If that is ever not    │
  │ true, switch HOLD_BUCKET_ENABLED to False below: everyone payable lands   │
  │ in the Pay sheet and the flags become review-only annotations.            │
  └───────────────────────────────────────────────────────────────────────────┘

FLAGGED vs CLEAN
────────────────
  The Hold bucket uses quality_filter's `exclude_recommended` (its own
  combination rule: >= exclude_min_flags independent flags, or all shown
  attention checks failed).  Single-flag-but-not-excluded respondents stay in
  the Pay bucket, but their n_flags and reasons are shown so you can eyeball
  them.  To split on "any flag >= 1" instead, set HOLD_ON_ANY_FLAG = True.

CLI
───
  python manage_payments.py            # update ledger + write report
  python manage_payments.py --dry-run  # print only; write nothing

Configuration
─────────────
  Reads ../config.yaml (the shared config).  Looks for the response export in
  the shared ingest/ folder; pass --export PATH to point at a specific file.
"""

from __future__ import annotations

import argparse
import csv
import io
import sys
import tempfile
import zipfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import yaml

# The vendored, canonical quality filter lives beside this file (kept in sync by
# sync_qc.bat).  Import it as a module so we reuse the exact same rule.
import quality_filter as qf


# ── POLICY SWITCHES ────────────────────────────────────────────────────────────
# See the header note. HOLD_BUCKET_ENABLED gates whether flagged completers are
# split out from payment at all.  HOLD_ON_ANY_FLAG changes the split criterion
# from "exclude_recommended" to "n_flags >= 1".
HOLD_BUCKET_ENABLED = True
# Moderate-stringency setting (June 2026): route ANY flagged completer (n_flags
# >= 1) to Hold for human review, instead of only those the quality filter
# recommends excluding (>= 2 flags). This changes the PAYMENT split only; the
# shared analysis exclusion rule in quality_config.json (exclude_min_flags) is
# deliberately left at 2 so the analysis repo's drop decision is unaffected.
HOLD_ON_ANY_FLAG = True

# Values in the hand-edited `paid` column that count as "already paid".
PAID_TRUE_VALUES = {"yes", "y", "true", "1", "paid", "done", "sent", "complete"}

# Columns in the payment ledger that the user edits by hand. Re-runs MUST
# preserve these — they are merged back by cid and never overwritten.
PRESERVED_COLS = ["paid", "paid_date", "paid_amount", "notes"]


# ── CONFIG & PATHS ──────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent                  # …/survey_management/payment_management/
REPO_ROOT = SCRIPT_DIR.parent                        # …/survey_management/
CONFIG_PATH = REPO_ROOT / "config.yaml"

with open(CONFIG_PATH, "r", encoding="utf-8") as _f:
    config = yaml.safe_load(_f)

INBOX_DIR = REPO_ROOT / config["paths"]["inbox_dir"]
TRACKER_PATH = REPO_ROOT / config["paths"]["tracker_file"]
PAYMENT_TRACKER_PATH = REPO_ROOT / "data" / "payment_tracker.xlsx"
REPORT_PATH = REPO_ROOT / "data" / "payment_report_unpaid.xlsx"

# Shared fraud blacklist, written by consent_management's suspicious-entry
# screen. One row per fraudulent consent Response ID + the IP it came from.
FRAUD_BLACKLIST_PATH = REPO_ROOT / "data" / "fraud_blacklist.csv"

# Manual "keep" decisions made in examine_indv's review queue. A reviewer who
# looks at a flagged completer and judges them legitimate records a cid here
# (decision=cleared). We honour that by forcing the cid OUT of the Hold bucket
# and into Pay, so a human keep survives every re-run of this script. It never
# overrides Fraud (fraud is decided before Hold/Pay split). Written by
# examine_indv/examine_review.py.
REVIEW_STATE_PATH = REPO_ROOT / "data" / "review_state.csv"

# Populated in main() from REVIEW_STATE_PATH; used by is_held().
KEPT_CIDS: set = set()

# Known throwaway / disposable email domains (public maintained list + curated
# additions) and a whitelist of real institution domains never to treat as
# disposable. A completer whose delivery email is on the throwaway list is
# routed to HOLD for review, not auto-paid — see is_held(). Both files are
# optional: if missing, the check simply never trips.
THROWAWAY_DOMAINS_PATH = REPO_ROOT / "data" / "email_throwaway_domains.txt"
EMAIL_WHITELIST_PATH = REPO_ROOT / "data" / "email_domain_whitelist.txt"

# Domain suffixes always treated as legitimate (schools / institutions), even if
# a domain somehow appears on the throwaway list. Protects the K-12 population.
WHITELIST_SUFFIXES = (".edu", ".k12.", ".sch.uk")

# The embedded-data column in the response export that carries the consent
# Response ID (= cid), and the respondent's IP column. Used to match a teen
# completion back to a blacklisted consent entry / IP.
RESPONSE_CID_FIELD = config.get("response_qualtrics_fields", {}).get("cid", "cid")
RESPONSE_IP_FIELD = "IPAddress"

CUTOFF_DATE = date.fromisoformat(config.get("response_cutoff_date", "2026-06-09"))


# ── HELPERS ──────────────────────────────────────────────────────────────────────

def _first_name(full_name: str) -> str:
    full_name = (full_name or "").strip()
    return full_name.split()[0] if full_name else ""


def _email_domain(email: str) -> str:
    """Lowercased domain part of an email, or '' if there isn't a clean one."""
    e = (email or "").strip().lower()
    if e.count("@") != 1:
        return ""
    return e.rsplit("@", 1)[1]


def _load_domain_file(path: Path) -> set:
    """Read a one-domain-per-line file, skipping blanks and # comments.
    Returns a lowercased set. Missing file -> empty set (check just won't trip)."""
    if not path.exists():
        return set()
    out = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip().lower()
        if line:
            out.add(line)
    return out


# Loaded once at import. Whitelist wins over the throwaway list.
THROWAWAY_DOMAINS = _load_domain_file(THROWAWAY_DOMAINS_PATH)
EMAIL_WHITELIST = _load_domain_file(EMAIL_WHITELIST_PATH)


def throwaway_reason(email: str) -> str:
    """If the delivery email's domain is a known throwaway (and not whitelisted),
    return a short human reason; otherwise ''. Whitelisted institution domains
    (.edu etc.) always return '' so real school addresses are never held."""
    dom = _email_domain(email)
    if not dom:
        return ""
    if dom in EMAIL_WHITELIST or dom.endswith(WHITELIST_SUFFIXES):
        return ""
    if dom in THROWAWAY_DOMAINS:
        return f"delivery email on throwaway-domain list ({dom})"
    return ""


def _is_paid(val) -> bool:
    return str(val).strip().lower() in PAID_TRUE_VALUES


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


def parse_emailed_at(ts_str: str):
    """Parse an emailed_at / completion timestamp to a date, or None."""
    if not ts_str or not str(ts_str).strip():
        return None
    s = str(ts_str).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def load_tracker() -> pd.DataFrame:
    if not TRACKER_PATH.exists():
        sys.exit(
            f"Tracker not found at {TRACKER_PATH}.\n"
            "Run consent_management/send_survey_emails.py and "
            "response_management/manage_responses.py first so completions exist."
        )
    return pd.read_excel(TRACKER_PATH, sheet_name="Participants", dtype=str).fillna("")


# A consent/screener export is told apart from the teen response export by a
# question stem that only exists on the consent survey. The inbox holds both
# types; payment wants the response export. Cheap heuristic over the header rows.
CONSENT_ONLY_MARKERS = (
    "i confirm that i have read",
    "i understand what the researchers asked",
    "parent/guardian full name",
)


def _peek_header_text(path: Path) -> str:
    """Lowercased text of the first 3 (header) rows of a .zip or .csv export."""
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".csv")]
            if not names:
                return ""
            grid = pd.read_csv(io.BytesIO(z.read(names[0])), dtype=str,
                               header=None, encoding="utf-8-sig", nrows=3)
    else:
        grid = pd.read_csv(path, dtype=str, header=None,
                           encoding="utf-8-sig", nrows=3)
    return " || ".join(" ".join(str(v).lower().split())
                       for v in grid.values.flatten())


def _looks_like_response_export(path: Path) -> bool:
    """True if headers are NOT a consent export (i.e. the teen response survey)."""
    try:
        text = _peek_header_text(path)
    except Exception:
        return False
    if not text.strip():
        return False
    return not any(m in text for m in CONSENT_ONLY_MARKERS)


def find_response_export(explicit: str | None):
    """Locate the response-survey export to feed the quality filter.

    With --export, use that path. Otherwise pick the most recently MODIFIED file
    in ingest/ (zip or csv) whose headers match the response-survey schema,
    quietly ignoring consent/screener exports. Returns a Path or None.
    """
    if explicit:
        p = Path(explicit)
        return p if p.exists() else None
    if not INBOX_DIR.exists():
        return None
    candidates = sorted(
        list(INBOX_DIR.glob("*.zip")) + list(INBOX_DIR.glob("*.csv")),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    for p in candidates:
        if _looks_like_response_export(p):
            return p
    return None


def quality_records_by_cid(export_path: Path) -> tuple[dict, str | None]:
    """Run quality_filter over the response export and return {cid: record}.

    quality_filter.load_qualtrics_csv reads the short-code header row (Q220,
    Duration, cid, ...), which is exactly what a Qualtrics export ZIP/CSV
    contains.  If the path is a ZIP we extract its first CSV to a temp file.

    Returns (records_by_cid, warning_or_None).
    """
    csv_path = export_path
    tmp = None
    if export_path.suffix.lower() == ".zip":
        with zipfile.ZipFile(export_path) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".csv")]
            if not names:
                return {}, f"No CSV inside {export_path.name}."
            tmp = tempfile.NamedTemporaryFile(
                suffix=".csv", delete=False, mode="wb")
            tmp.write(z.read(names[0]))
            tmp.close()
            csv_path = Path(tmp.name)

    try:
        records, _cfg = qf.evaluate_export(str(csv_path))
    except Exception as e:
        return {}, f"Could not evaluate {export_path.name}: {e}"
    finally:
        if tmp is not None:
            try:
                Path(tmp.name).unlink()
            except OSError:
                pass

    by_cid = {}
    for r in records:
        # quality_filter.respondent_id prefers cid; fall back to response_id.
        key = str(r.get("id", "")).strip() or str(r.get("response_id", "")).strip()
        if key:
            by_cid[key] = r
    return by_cid, None


# ── FRAUD BLACKLIST ─────────────────────────────────────────────────────────

def load_fraud_blacklist(path: Path = FRAUD_BLACKLIST_PATH) -> tuple[set, set, dict]:
    """Read data/fraud_blacklist.csv (written by the consent screen).

    Returns (fraud_ids, fraud_ips, reason_by_key) where reason_by_key is keyed
    by both response_id and ip so we can report WHY a row was flagged. Missing
    file is fine — it just means nothing is blacklisted yet.
    """
    fraud_ids, fraud_ips, reasons = set(), set(), {}
    if not path.exists():
        return fraud_ids, fraud_ips, reasons
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            rid = str(row.get("response_id", "")).strip()
            ip = str(row.get("ip", "")).strip()
            why = str(row.get("reasons", "")).strip()
            if rid:
                fraud_ids.add(rid)
                reasons.setdefault(rid, why)
            if ip:
                fraud_ips.add(ip)
                reasons.setdefault(ip, why)
    return fraud_ids, fraud_ips, reasons


def response_ip_by_cid(export_path: Path | None) -> dict:
    """Map each completion's cid -> the IP it was submitted from, read straight
    from the response export. Lets payment match a completion against
    blacklisted IPs (not just blacklisted cids). Empty if no export/columns."""
    if export_path is None:
        return {}
    csv_path, tmp = export_path, None
    try:
        if export_path.suffix.lower() == ".zip":
            with zipfile.ZipFile(export_path) as z:
                names = [n for n in z.namelist() if n.lower().endswith(".csv")]
                if not names:
                    return {}
                tmp = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb")
                tmp.write(z.read(names[0]))
                tmp.close()
                csv_path = Path(tmp.name)
        codes, rows = qf.load_qualtrics_csv(str(csv_path))
    except Exception:
        return {}
    finally:
        if tmp is not None:
            try:
                Path(tmp.name).unlink()
            except OSError:
                pass

    if RESPONSE_CID_FIELD not in codes or RESPONSE_IP_FIELD not in codes:
        return {}
    out = {}
    for r in rows:
        cid = str(r.get(RESPONSE_CID_FIELD, "")).strip()
        ip = str(r.get(RESPONSE_IP_FIELD, "")).strip()
        if cid and cid not in out:   # keep the first (earliest) completion's IP
            out[cid] = ip
    return out


# ── CORE ──────────────────────────────────────────────────────────────────────

def build_payment_rows(tracker_df: pd.DataFrame, quality: dict,
                       fraud_ids: set | None = None,
                       fraud_ips: set | None = None,
                       fraud_reasons: dict | None = None,
                       ip_by_cid: dict | None = None) -> list[dict]:
    """One row per completed participant, with consent/completion confirmed,
    quality flags merged on cid, and a fraud verdict from the blacklist."""
    fraud_ids = fraud_ids or set()
    fraud_ips = fraud_ips or set()
    fraud_reasons = fraud_reasons or {}
    ip_by_cid = ip_by_cid or {}

    rows = []
    for _, t in tracker_df.iterrows():
        completed_at = str(t.get("survey_completed_at", "")).strip()
        if not completed_at:
            continue  # not completed → not payable

        cid = str(t.get("response_id", "")).strip()

        # Respect the response cutoff: skip stale pre-pilot completions.
        cdate = parse_emailed_at(completed_at)
        if cdate is not None and cdate < CUTOFF_DATE:
            continue

        child = str(t.get("child_name", "")).strip()
        q = quality.get(cid)

        if q is None:
            quality_status = "NOT EVALUATED"
            n_flags = ""
            exclude = False
            reasons = "response export missing or cid not found in export"
        else:
            quality_status = "evaluated"
            n_flags = q.get("n_flags", "")
            exclude = bool(q.get("exclude_recommended", False))
            reasons = q.get("reasons", "")

        # Fraud check: the consent cid is blacklisted, or the survey was taken
        # from a blacklisted IP. Either is enough — fraud is never paid.
        survey_ip = str(ip_by_cid.get(cid, "")).strip()
        fraud_why = []
        if cid in fraud_ids:
            fraud_why.append(f"cid blacklisted ({fraud_reasons.get(cid, 'flagged at consent')})")
        if survey_ip and survey_ip in fraud_ips:
            fraud_why.append(f"survey IP {survey_ip} blacklisted")
        is_fraud_row = bool(fraud_why)

        # Don't send a gift card to a known throwaway / disposable email. This
        # routes the row to HOLD for review (see is_held), never auto-pay.
        delivery_email = str(t.get("delivery_email", "")).strip()
        tw_reason = throwaway_reason(delivery_email)
        if tw_reason:
            reasons = f"{reasons}; {tw_reason}".strip("; ") if reasons else tw_reason

        rows.append({
            "cid": cid,
            "first_name": _first_name(child),
            "child_name": child,
            "delivery_email": delivery_email,
            "survey_ip": survey_ip,
            "survey_completed_at": completed_at,
            "consent_ok": "yes",                 # implied by presence in tracker
            "completed": "yes",
            "quality_status": quality_status,
            "n_flags": n_flags,
            "exclude_recommended": "yes" if exclude else "no",
            "flag_reasons": reasons,
            "throwaway_email": "yes" if tw_reason else "no",
            "fraud": "yes" if is_fraud_row else "no",
            "fraud_reason": "; ".join(fraud_why),
        })
    return rows


def is_fraud(row: dict) -> bool:
    """Whether this payable row is blacklisted as fraud (never paid)."""
    return str(row.get("fraud", "no")).strip().lower() == "yes"


def load_kept_cids() -> set:
    """Read review_state.csv -> set of cids a reviewer explicitly kept.

    Later rows win, so a cid can be cleared then later un-cleared. Tolerant of a
    missing/half-written file (returns what it can, never raises)."""
    kept: set = set()
    if not REVIEW_STATE_PATH.exists():
        return kept
    try:
        import csv as _csv
        with open(REVIEW_STATE_PATH, newline="", encoding="utf-8-sig") as f:
            for r in _csv.DictReader(f):
                cid = str(r.get("cid", "")).strip()
                if not cid:
                    continue
                if str(r.get("decision", "")).strip().lower() == "cleared":
                    kept.add(cid)
                else:
                    kept.discard(cid)
    except OSError:
        pass
    return kept


def is_held(row: dict) -> bool:
    """Whether this payable row belongs in the Hold bucket."""
    # A reviewer who explicitly KEPT this completer in examine_indv overrides
    # every hold rule below (including throwaway-email): the human looked and
    # judged them legitimate, so they go to Pay and stay there on re-runs.
    if str(row.get("cid", "")).strip() in KEPT_CIDS:
        return False
    # A throwaway delivery email always holds for review, independent of the
    # quality-flag switches below: we never auto-send a card to one.
    if str(row.get("throwaway_email", "no")).strip().lower() == "yes":
        return True
    if not HOLD_BUCKET_ENABLED:
        return False
    if HOLD_ON_ANY_FLAG:
        try:
            return int(row.get("n_flags") or 0) >= 1
        except (TypeError, ValueError):
            return False
    return str(row.get("exclude_recommended", "no")).strip().lower() == "yes"


def upsert_ledger(payment_rows: list[dict]) -> pd.DataFrame:
    """Merge the freshly-built payment rows into the persistent ledger,
    preserving hand-edited columns (paid, paid_date, paid_amount, notes) by cid.
    """
    new_df = pd.DataFrame(payment_rows)
    for col in PRESERVED_COLS:
        if col not in new_df.columns:
            new_df[col] = ""

    if PAYMENT_TRACKER_PATH.exists():
        old = pd.read_excel(PAYMENT_TRACKER_PATH, sheet_name="Payments",
                            dtype=str).fillna("")
        old_by_cid = {str(r["cid"]).strip(): r for _, r in old.iterrows()}
        for i, r in new_df.iterrows():
            prev = old_by_cid.get(str(r["cid"]).strip())
            if prev is not None:
                for col in PRESERVED_COLS:
                    new_df.at[i, col] = str(prev.get(col, "")).strip()
    return new_df


def save_ledger(df: pd.DataFrame):
    PAYMENT_TRACKER_PATH.parent.mkdir(parents=True, exist_ok=True)
    col_order = [
        "cid", "first_name", "child_name", "delivery_email", "survey_ip",
        "survey_completed_at", "consent_ok", "completed",
        "quality_status", "n_flags", "exclude_recommended", "flag_reasons",
        "throwaway_email", "fraud", "fraud_reason",
        "paid", "paid_date", "paid_amount", "notes",
    ]
    cols = [c for c in col_order if c in df.columns] + \
           [c for c in df.columns if c not in col_order]
    with pd.ExcelWriter(PAYMENT_TRACKER_PATH, engine="openpyxl") as xl:
        df[cols].to_excel(xl, sheet_name="Payments", index=False)


def bucket_rows(payment_rows: list[dict]) -> tuple[list, list, list]:
    """Split UNPAID completers into (pay, hold, fraud). Fraud takes priority:
    a blacklisted row never lands in Pay or Hold regardless of quality flags."""
    unpaid = [r for r in payment_rows if not _is_paid(r.get("paid", ""))]
    fraud_rows = [r for r in unpaid if is_fraud(r)]
    rest = [r for r in unpaid if not is_fraud(r)]
    pay_rows = [r for r in rest if not is_held(r)]
    hold_rows = [r for r in rest if is_held(r)]
    return pay_rows, hold_rows, fraud_rows


def write_report(payment_rows: list[dict]):
    """Write the unpaid report: Pay, Hold, and Fraud sheets."""
    report_cols = ["first_name", "delivery_email", "cid",
                   "n_flags", "flag_reasons", "survey_completed_at"]
    fraud_cols = ["first_name", "delivery_email", "cid", "survey_ip",
                  "fraud_reason", "survey_completed_at"]

    pay_rows, hold_rows, fraud_rows = bucket_rows(payment_rows)

    def _frame(rows, cols):
        if not rows:
            return pd.DataFrame(columns=cols)
        df = pd.DataFrame(rows)[cols]
        return df.sort_values("first_name", key=lambda s: s.str.lower())

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as xl:
        _frame(pay_rows, report_cols).to_excel(xl, sheet_name="Pay (no hold)", index=False)
        _frame(hold_rows, report_cols).to_excel(xl, sheet_name="Hold (flagged)", index=False)
        _frame(fraud_rows, fraud_cols).to_excel(xl, sheet_name="Fraud (do not pay)", index=False)
    return pay_rows, hold_rows, fraud_rows


# ── MAIN ────────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Payment management: confirm consent+completion, apply the "
                    "quality filter, track who's paid, and export the unpaid "
                    "Amazon-card report.")
    ap.add_argument("--export", help="Path to the response-survey export "
                                      "(ZIP or CSV). Defaults to newest in ingest/.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Print only — write no ledger or report files.")
    args = ap.parse_args()

    # Fail fast: the files we write at the end must be writable now, so we never
    # do the whole quality+fraud build and then crash on a locked output file.
    # (Dry-run writes nothing, so it skips this check.)
    if not args.dry_run:
        try:
            _assert_writable(PAYMENT_TRACKER_PATH, "the payment ledger")
            _assert_writable(REPORT_PATH, "the unpaid report")
        except Exception as e:
            print(f"PREFLIGHT FAILED: {e}\n")
            return

    print(f"\nTracker        : {TRACKER_PATH}")
    print(f"Payment ledger : {PAYMENT_TRACKER_PATH}")
    print(f"Cutoff date    : {CUTOFF_DATE}")
    print(f"Hold bucket    : {'ON' if HOLD_BUCKET_ENABLED else 'OFF'} "
          f"({'any flag' if HOLD_ON_ANY_FLAG else 'exclude_recommended'})")

    fraud_ids, fraud_ips, fraud_reasons = load_fraud_blacklist()
    if fraud_ids or fraud_ips:
        print(f"Fraud blacklist: {len(fraud_ids)} response ID(s), "
              f"{len(fraud_ips)} IP(s)  ({FRAUD_BLACKLIST_PATH.name})\n")
    else:
        print(f"Fraud blacklist: none found at {FRAUD_BLACKLIST_PATH.name} "
              f"(no fraud screening yet)\n")

    # Honour manual "keep" decisions from examine_indv's review queue: these
    # cids are forced out of Hold into Pay (see is_held / load_kept_cids).
    global KEPT_CIDS
    KEPT_CIDS = load_kept_cids()
    if KEPT_CIDS:
        print(f"Reviewer keeps : {len(KEPT_CIDS)} cid(s) kept in Pay "
              f"({REVIEW_STATE_PATH.name})\n")

    tracker_df = load_tracker()

    export_path = find_response_export(args.export)
    if export_path is None:
        print("WARNING: No response-survey export found in ingest/ (or via "
              "--export).\n  Completions will still be read from the tracker, "
              "but quality flags cannot be computed, so every payable\n  "
              "participant lands in the Pay sheet marked 'NOT EVALUATED'.\n")
        quality = {}
    else:
        print(f"Quality source : {export_path.name}")
        quality, warn = quality_records_by_cid(export_path)
        if warn:
            print(f"  WARNING: {warn}\n  Proceeding without quality flags.\n")
        else:
            print(f"  Evaluated {len(quality)} respondent(s) from the export.\n")

    # Map each completion to the IP it was taken from, so a completion from a
    # blacklisted IP is caught even when its cid itself isn't on the list.
    ip_by_cid = response_ip_by_cid(export_path)

    payment_rows = build_payment_rows(tracker_df, quality, fraud_ids, fraud_ips,
                                      fraud_reasons, ip_by_cid)
    if not payment_rows:
        print("No completed participants found in the tracker. Nothing to pay.\n")
        return

    ledger = upsert_ledger(payment_rows)
    # Carry preserved fields back onto the in-memory rows so the report respects
    # the `paid` column even on the first build.
    paid_by_cid = {str(r["cid"]).strip(): r for _, r in ledger.iterrows()}
    for r in payment_rows:
        prev = paid_by_cid.get(str(r["cid"]).strip())
        if prev is not None:
            for col in PRESERVED_COLS:
                r[col] = str(prev.get(col, "")).strip()

    pay_rows, hold_rows, fraud_rows = bucket_rows(payment_rows)

    n_completed = len(payment_rows)
    n_paid = sum(1 for r in payment_rows if _is_paid(r.get("paid", "")))
    # Anyone fraud-matched who is ALSO already marked paid: surface loudly so
    # you know a card may have gone to a scammer.
    fraud_already_paid = [r for r in payment_rows
                          if is_fraud(r) and _is_paid(r.get("paid", ""))]

    print("═" * 60)
    print("  PAYMENT SUMMARY")
    print("═" * 60)
    print(f"  Completed (consent + survey) : {n_completed}")
    print(f"  Already marked paid          : {n_paid}")
    print(f"  Unpaid → Pay (no hold)       : {len(pay_rows)}")
    print(f"  Unpaid → Hold (flagged)      : {len(hold_rows)}")
    print(f"  Unpaid → Fraud (DO NOT PAY)  : {len(fraud_rows)}")
    print()
    if pay_rows:
        print("  PAY — buy + send Amazon cards to:")
        for r in sorted(pay_rows, key=lambda x: x['first_name'].lower()):
            print(f"    • {r['first_name']:<15} <{r['delivery_email']}>")
        print()
    if hold_rows:
        print("  HOLD — review quality flags before deciding:")
        for r in sorted(hold_rows, key=lambda x: x['first_name'].lower()):
            print(f"    • {r['first_name']:<15} <{r['delivery_email']}>  "
                  f"[{r['flag_reasons']}]")
        print()
    if fraud_rows:
        print("  FRAUD — blacklisted, DO NOT pay:")
        for r in sorted(fraud_rows, key=lambda x: x['first_name'].lower()):
            print(f"    • {r['first_name']:<15} <{r['delivery_email']}>  "
                  f"[{r['fraud_reason']}]")
        print()
    if fraud_already_paid:
        print("  ⚠ WARNING — these are blacklisted but ALREADY marked paid:")
        for r in fraud_already_paid:
            print(f"    • {r['first_name']:<15} <{r['delivery_email']}>  "
                  f"[{r['fraud_reason']}]")
        print()

    if args.dry_run:
        print("Dry run — no files written.\n")
        return

    save_ledger(ledger)
    write_report(payment_rows)
    print(f"  Ledger written : {PAYMENT_TRACKER_PATH}")
    print(f"  Report written : {REPORT_PATH}")
    print("  (Edit the `paid`/`paid_date` columns in the ledger as you send "
          "cards;\n   re-runs preserve those edits.)\n")


if __name__ == "__main__":
    main()
