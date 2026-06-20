"""
flag_suspicious_entries.py
──────────────────────────
Flag consent/screener submissions that look like bots, fraud farms, or
duplicate/automated entries, BEFORE send_survey_emails.py invites them.

WHY THIS IS SEPARATE FROM quality_filter.py
    quality_filter.py judges the *teen survey* responses (attention checks,
    page-read times, straight-lining). This module judges the *consent export*,
    which has different fraud signals: many submissions from one IP, the same
    email + child submitted twice, throwaway-fast completions, low reCAPTCHA
    scores, and one person signing both the parent and child lines.

WHAT IT REUSES
    The ZIP reading and column resolution are imported from send_survey_emails
    so the two scripts can never drift on how they parse a Qualtrics export:
        load_csv_from_zip, resolve_columns, _norm

WHAT IT DELIBERATELY DOES NOT FLAG
    - A parent who submits twice with DIFFERENT children (siblings) is normal.
      Only same email + same child counts as a duplicate submission.
    - A signature that differs from the typed name only by a middle name,
      initial, or spacing. Real people do this constantly. We only flag when
      the parent and child signatures are byte-for-byte the same person.
    - A household sharing one IP for 2-3 kids. The IP-farm flag needs a larger
      cluster; small shared-IP counts only matter combined with another signal.
    - A name that "looks" unusual. There is NO linguistic / dictionary check, so
      non-Western names are never penalised. The one name signal (name_swap) is
      purely structural — child surname == parent given name — and even that is
      only a CONTEXTUAL soft flag: it can never mark an entry suspicious on its
      own, because that same pattern is normal in patronymic naming (e.g. South
      India). It only escalates an entry that already has a technical signal.

OUTPUT
    flag_entries() returns one record dict per row with independent boolean
    flags, a list of human-readable reasons, and a `suspicious` verdict. It
    decides nothing about emailing; the caller applies policy.

CLI
    python flag_suspicious_entries.py [path/to/export.zip] [-o report.csv]
    With no path, the most recent consent ZIP in ../ingest is used.
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

# Reuse the exact ZIP/column logic the mailer uses, so parsing never diverges.
from send_survey_emails import (
    load_csv_from_zip,
    resolve_columns,
    _norm,
    INBOX_DIR,
    TRACKER_PATH,
    pick_consent_zip,
    correct_email_typo,
    email_domain,
)

# The fraud blacklist is the shared contract between this screen and payment
# processing. It lives in the same data/ folder as the tracker and logs, so
# both consent_management and payment_management resolve the same file.
FRAUD_BLACKLIST_PATH = TRACKER_PATH.parent / "fraud_blacklist.csv"
BLACKLIST_COLS = ["response_id", "ip", "child_name", "delivery_email",
                  "reasons", "source", "added_at"]

# Known throwaway / disposable email domains + an institution whitelist, both in
# data/ beside the tracker. A delivery email whose domain is on the list (and not
# whitelisted) raises the throwaway_email TECHNICAL soft flag. Missing files just
# mean the flag never trips.
THROWAWAY_DOMAINS_PATH = TRACKER_PATH.parent / "email_throwaway_domains.txt"
EMAIL_WHITELIST_PATH = TRACKER_PATH.parent / "email_domain_whitelist.txt"
WHITELIST_SUFFIXES = (".edu", ".k12.", ".sch.uk")


def _load_domain_set(path) -> set:
    if not path.exists():
        return set()
    out = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip().lower()
        if line:
            out.add(line)
    return out


THROWAWAY_DOMAINS = _load_domain_set(THROWAWAY_DOMAINS_PATH)
EMAIL_WHITELIST = _load_domain_set(EMAIL_WHITELIST_PATH)


def is_throwaway_email(email) -> bool:
    """True if the email's domain is a known throwaway and not whitelisted.
    Typo correction runs first, so a slip like 'gmil.com' is fixed to gmail.com
    and never counts as throwaway."""
    dom = email_domain(email)
    if not dom or dom in EMAIL_WHITELIST or dom.endswith(WHITELIST_SUFFIXES):
        return False
    return dom in THROWAWAY_DOMAINS


# ── THRESHOLDS ──────────────────────────────────────────────────────────────
# Conservative on purpose: it is worse to wrongly flag a real family than to
# let one borderline entry through to manual review. Tune on real data and
# record changes; these were set against the June pilot exports.
DEFAULT_THRESHOLDS = {
    # One IP submitting at least this many finished consents = a farm. Real
    # households top out around 2-4 siblings, so 5 clears them comfortably.
    "ip_farm_min": 5,
    # Weaker signal: one IP used by at least this many DISTINCT emails. Only
    # contributes toward the 2-flag combination rule, never alone.
    "ip_shared_emails_min": 3,
    # Qualtrics reCAPTCHA v3 score; below this is bot-like. Raised 0.5 -> 0.7 to
    # match the teen-survey moderate setting. ~18% of consents score below 0.7,
    # but this stays a SOFT flag, so a low score alone never blacklists.
    "recaptcha_min": 0.7,
    # SOFT speeding floor: finishing faster than this many seconds is suspicious,
    # but only blacklists in combination with a second signal. Set to 76s, the
    # lower-1-sigma "suspicious-fast" floor from the OUTLIER-EXCLUDED consent
    # finish-time distribution. (Raw mean/SD are unusable: one 22.5h entry gives
    # CV=8.4 and mean-k*SD goes negative. After Tukey IQR-1.5x trimming the mean
    # is 157.9s / SD 82.5s -> mean-1SD=75s; a log-normal fit gives
    # exp(mu-1sigma)=76s. The two agree, so 76s is the data-derived floor.)
    "fast_seconds": 76,
    # HARD speeding floor: finishing faster than this is physically implausible
    # for a form requiring reading + two typed signatures + an email, so it
    # blacklists on its own. Set below the fastest real completion observed (46s)
    # so it catches future ultra-fast bots with ~0 false positives today.
    "fast_hard_seconds": 30,
    # Number of independent TECHNICAL soft flags (ip_shared / low_recaptcha /
    # fast) needed to call an entry suspicious on their own. A single technical
    # flag plus a contextual one (same_signature / name_swap) also suffices.
    "tech_soft_to_flag": 2,
}

# Metadata columns are not in send_survey_emails.FIELD_MATCHERS, so resolve
# them here by a distinctive substring of the human-readable label row.
META_MATCHERS = {
    "ip":        "ip address",
    "recaptcha": "recaptchascore",
    "duration":  "duration (in seconds)",
}


def _resolve_meta(df: pd.DataFrame) -> dict:
    """Map metadata fields (ip, recaptcha, duration) to actual column names.
    Missing metadata is tolerated: the corresponding flags just never trip."""
    norm_cols = [(_norm(c), c) for c in df.columns]
    out = {}
    for field, needle in META_MATCHERS.items():
        match = next((orig for n, orig in norm_cols if needle in n), None)
        if match is not None:
            out[field] = match
    return out


def _to_float(val):
    try:
        return float(str(val).strip())
    except (TypeError, ValueError):
        return None


def _finished(val) -> bool:
    return str(val).strip().upper() in ("TRUE", "1", "YES")


def _name_token_swap(parent_name: str, child_name: str) -> bool:
    """Language-neutral structural signal of a fabricated name pair.

    In a real family the child shares the parent's SURNAME in the SAME position
    (parent 'Iris Fels' -> child 'Sabelle Fels'). Fraud farms recycle a small
    pool of tokens and drop them in the wrong slot, so the child's *surname*
    ends up being the parent's *given* name (or vice-versa):
        'Gamey locale' / 'Doole gamey'   ('gamey' = parent first, child last)
        'Jones gadey'  / 'Gahay jones'
        'Hande readde' / 'Defau hande'

    Returns True only for that cross-position reuse. It deliberately does NOT
    judge whether a name 'looks real', so it never penalises non-Western names;
    a normal shared surname, a different surname, or a single-token name all
    return False.

    CAVEAT: a real family that writes the family name first inconsistently
    (parent 'Smith Jane', child 'Bob Smith') also matches. That is why this is
    only a SOFT flag — it must be corroborated by another signal to mark an
    entry suspicious.
    """
    p = [t for t in str(parent_name).lower().split() if t]
    c = [t for t in str(child_name).lower().split() if t]
    if len(p) < 2 or len(c) < 2:
        return False
    pf, pl = p[0], p[-1]
    cf, cl = c[0], c[-1]
    if cl == pl:          # normal shared surname (same position) — not a swap
        return False
    return cl == pf or cf == pl


def flag_entries(df: pd.DataFrame, thresholds: dict | None = None) -> list[dict]:
    """Score every row of a consent export for bot/fraud/duplicate signals.

    Args:
        df: DataFrame from load_csv_from_zip (human-readable label header row).
        thresholds: optional overrides for DEFAULT_THRESHOLDS.

    Returns:
        A list of record dicts, one per FINISHED response, each with:
          response_id, delivery_email, ip, recorded_date,
          flag_ip_farm, flag_dup_submission, flag_ip_shared,
          flag_low_recaptcha, flag_fast, flag_same_signature,
          n_hard_flags, n_soft_flags, suspicious (bool), reasons (str).
        Incomplete responses are skipped (the mailer already excludes them).
    """
    cfg = dict(DEFAULT_THRESHOLDS)
    if thresholds:
        cfg.update(thresholds)

    cols = resolve_columns(df)          # response_id, parent/child name+sig, email...
    meta = _resolve_meta(df)            # ip, recaptcha, duration (any may be absent)

    c_id = cols["response_id"]
    c_email = cols["delivery_email"]
    c_child = cols["child_name"]
    c_child_name = cols["child_name"]
    c_parent = cols["parent_name"]
    c_psig = cols["parent_sig"]
    c_csig = cols["child_sig"]
    c_finished = cols["finished"]
    c_recorded = cols.get("recorded_date")
    c_ip = meta.get("ip")
    c_rc = meta.get("recaptcha")
    c_dur = meta.get("duration")

    # Only finished responses can be fraud worth flagging; the rest are noise.
    rows = [r for _, r in df.iterrows() if _finished(r.get(c_finished, ""))]

    # ── Batch-level indexes ──────────────────────────────────────────────────
    ip_row_count = Counter()
    ip_distinct_emails = defaultdict(set)
    email_child_count = Counter()       # (email, child) -> times submitted
    for r in rows:
        ip = str(r.get(c_ip, "")).strip() if c_ip else ""
        email = _norm(correct_email_typo(r.get(c_email, ""))[0])
        child = _norm(r.get(c_child, ""))
        if ip:
            ip_row_count[ip] += 1
            if email:
                ip_distinct_emails[ip].add(email)
        if email:
            email_child_count[(email, child)] += 1

    records = []
    for r in rows:
        rid = str(r.get(c_id, "")).strip()
        ip = str(r.get(c_ip, "")).strip() if c_ip else ""
        email = _norm(correct_email_typo(r.get(c_email, ""))[0])
        child = _norm(r.get(c_child, ""))
        raw_email = str(r.get(c_email, "")).strip()
        corrected_email, email_fixed = correct_email_typo(raw_email)
        psig = _norm(r.get(c_psig, ""))
        csig = _norm(r.get(c_csig, ""))
        rc = _to_float(r.get(c_rc, "")) if c_rc else None
        dur = _to_float(r.get(c_dur, "")) if c_dur else None

        # ── HARD flags: each alone is enough to mark the entry suspicious ─────
        flag_ip_farm = bool(ip and ip_row_count[ip] >= cfg["ip_farm_min"])
        flag_dup_submission = bool(email and email_child_count[(email, child)] > 1)

        # ── SOFT flags: need two together (or one alongside a hard flag) ──────
        flag_ip_shared = bool(
            ip
            and len(ip_distinct_emails[ip]) >= cfg["ip_shared_emails_min"]
            and ip_row_count[ip] >= cfg["ip_shared_emails_min"]
        )
        flag_low_recaptcha = bool(rc is not None and rc < cfg["recaptcha_min"])
        # Two-tier speeding: HARD = implausibly fast (blacklists alone); SOFT =
        # suspiciously fast band [fast_hard_seconds, fast_seconds) that needs
        # corroboration. The band keeps the two tiers from double-reporting.
        flag_fast_hard = bool(dur is not None and dur < cfg["fast_hard_seconds"])
        flag_fast = bool(dur is not None
                         and cfg["fast_hard_seconds"] <= dur < cfg["fast_seconds"])
        flag_throwaway_email = is_throwaway_email(corrected_email)
        flag_same_signature = bool(psig and csig and psig == csig)
        flag_name_swap = _name_token_swap(r.get(c_parent, ""), r.get(c_child_name, ""))

        # ── HARD: each alone is enough. ──────────────────────────────────────
        hard = [
            ("ip_farm", flag_ip_farm, f"{ip_row_count[ip]} entries from one IP" if ip else ""),
            ("dup_submission", flag_dup_submission, "same email + child submitted more than once"),
            ("fast_hard", flag_fast_hard,
             f"finished in {int(dur)}s (implausibly fast)" if dur is not None else ""),
        ]
        # ── TECHNICAL soft: network / automation signals. Two of these, or one
        #    of these plus a contextual signal, marks an entry suspicious. ─────
        tech_soft = [
            ("ip_shared", flag_ip_shared, f"IP shared by {len(ip_distinct_emails[ip])} emails" if ip else ""),
            ("low_recaptcha", flag_low_recaptcha, f"reCAPTCHA {rc:.2f}" if rc is not None else ""),
            ("fast", flag_fast, f"finished in {int(dur)}s" if dur is not None else ""),
            ("throwaway_email", flag_throwaway_email,
             f"throwaway email domain ({email_domain(corrected_email)})" if flag_throwaway_email else ""),
        ]
        # ── CONTEXTUAL soft: signals with plausible benign explanations
        #    (a parent signing both lines; patronymic / family-name-first naming
        #    that is normal in e.g. South India). These NEVER fire alone and
        #    NEVER combine with each other — they only escalate an entry that
        #    already has a technical signal. This keeps the name check from
        #    auto-excluding a real non-Western family. ─────────────────────────
        contextual_soft = [
            ("same_signature", flag_same_signature, "parent and child signatures identical"),
            ("name_swap", flag_name_swap,
             "child surname = parent's given name (possible fabricated name pair)"),
        ]

        n_hard = sum(1 for _, ok, _ in hard if ok)
        n_tech = sum(1 for _, ok, _ in tech_soft if ok)
        n_contextual = sum(1 for _, ok, _ in contextual_soft if ok)

        suspicious = (
            n_hard >= 1
            or n_tech >= cfg["tech_soft_to_flag"]
            or (n_tech >= 1 and n_contextual >= 1)
        )

        reasons = [msg for _, ok, msg in hard + tech_soft + contextual_soft if ok and msg]

        records.append({
            "response_id": rid,
            "child_name": str(r.get(c_child_name, "")).strip(),
            "delivery_email": corrected_email,
            "delivery_email_original": raw_email,
            "email_corrected": email_fixed,
            "ip": ip,
            "recorded_date": str(r.get(c_recorded, "")).strip() if c_recorded else "",
            "flag_ip_farm": flag_ip_farm,
            "flag_dup_submission": flag_dup_submission,
            "flag_ip_shared": flag_ip_shared,
            "flag_low_recaptcha": flag_low_recaptcha,
            "flag_fast": flag_fast,
            "flag_fast_hard": flag_fast_hard,
            "flag_throwaway_email": flag_throwaway_email,
            "flag_same_signature": flag_same_signature,
            "flag_name_swap": flag_name_swap,
            "n_hard_flags": n_hard,
            "n_soft_flags": n_tech + n_contextual,
            "suspicious": suspicious,
            "reasons": "; ".join(reasons),
        })

    return records


def suspicious_ids(records: list[dict]) -> set:
    """Convenience: the set of response IDs flagged suspicious."""
    return {r["response_id"] for r in records if r["suspicious"] and r["response_id"]}


def flag_zip(zip_path, thresholds: dict | None = None) -> list[dict]:
    """Read a consent ZIP and flag its entries in one call."""
    df = load_csv_from_zip(str(zip_path))
    return flag_entries(df, thresholds)


# ── FRAUD BLACKLIST (shared with payment_management) ─────────────────────────
# A persistent, append-only list of fraudulent consent Response IDs and the IPs
# they came from. The consent screen writes it; payment processing reads it to
# refuse a gift card to anyone whose cid or IP is on the list (a Fraud bucket,
# separate from the quality Hold bucket). The file is plain CSV so you can open
# it and add or remove rows by hand; manual rows are preserved across re-runs.


def load_blacklist(path=None) -> list[dict]:
    """Return the blacklist as a list of row dicts (empty if the file is absent)."""
    path = Path(path) if path else FRAUD_BLACKLIST_PATH
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return [dict(r) for r in csv.DictReader(fh)]


def load_blacklist_sets(path=None) -> tuple[set, set]:
    """Return (fraud_response_ids, fraud_ips) for fast membership checks."""
    ids, ips = set(), set()
    for row in load_blacklist(path):
        rid = str(row.get("response_id", "")).strip()
        ip = str(row.get("ip", "")).strip()
        if rid:
            ids.add(rid)
        if ip:
            ips.add(ip)
    return ids, ips


def write_blacklist(records: list[dict], path=None) -> int:
    """Upsert suspicious entries into the blacklist, preserving existing rows.

    `records` are flag_entries() records; only the suspicious ones are written.
    Rows already present (matched by response_id) keep their original source and
    added_at, so hand-added rows (source='manual') and earlier auto rows are
    never clobbered or dropped. Returns the number of NEW rows added.
    """
    path = Path(path) if path else FRAUD_BLACKLIST_PATH
    existing = {str(r.get("response_id", "")).strip(): r
                for r in load_blacklist(path) if str(r.get("response_id", "")).strip()}
    now = datetime.now().isoformat(timespec="seconds")

    added = 0
    for rec in records:
        if not rec.get("suspicious"):
            continue
        rid = str(rec.get("response_id", "")).strip()
        if not rid:
            continue
        if rid in existing:
            # Refresh the detection details but keep provenance fields intact.
            row = existing[rid]
            row["ip"] = rec.get("ip", row.get("ip", ""))
            row["child_name"] = rec.get("child_name", row.get("child_name", ""))
            row["delivery_email"] = rec.get("delivery_email", row.get("delivery_email", ""))
            row["reasons"] = rec.get("reasons", row.get("reasons", ""))
        else:
            existing[rid] = {
                "response_id": rid,
                "ip": rec.get("ip", ""),
                "child_name": rec.get("child_name", ""),
                "delivery_email": rec.get("delivery_email", ""),
                "reasons": rec.get("reasons", ""),
                "source": "auto",
                "added_at": now,
            }
            added += 1

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=BLACKLIST_COLS)
        w.writeheader()
        for row in existing.values():
            w.writerow({k: row.get(k, "") for k in BLACKLIST_COLS})
    return added


# ── Reporting / CLI ──────────────────────────────────────────────────────────
_FIELDNAMES = [
    "response_id", "delivery_email", "ip", "recorded_date",
    "flag_ip_farm", "flag_dup_submission", "flag_ip_shared",
    "flag_low_recaptcha", "flag_fast", "flag_fast_hard",
    "flag_same_signature", "flag_name_swap",
    "n_hard_flags", "n_soft_flags", "suspicious", "reasons",
]


def write_report(records: list[dict], out_path) -> None:
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=_FIELDNAMES)
        w.writeheader()
        for r in records:
            w.writerow({k: r.get(k, "") for k in _FIELDNAMES})


def print_summary(records: list[dict]) -> None:
    flagged = [r for r in records if r["suspicious"]]
    reason_counts = Counter()
    for r in flagged:
        for key in ("ip_farm", "dup_submission", "fast_hard", "ip_shared",
                    "low_recaptcha", "fast", "same_signature", "name_swap"):
            if r[f"flag_{key}"]:
                reason_counts[key] += 1
    print("=" * 64)
    print("SUSPICIOUS-ENTRY REPORT (consent/screener export)")
    print("=" * 64)
    print(f"Finished responses : {len(records)}")
    print(f"Flagged suspicious : {len(flagged)}")
    for key, n in reason_counts.most_common():
        print(f"    {key:<16}: {n}")
    print("-" * 64)
    for r in flagged:
        print(f"  {r['delivery_email'] or '(no email)':<34} {r['ip']:<16} {r['reasons']}")
    print()


def _default_zip():
    zips = sorted(Path(INBOX_DIR).glob("*.zip")) if Path(INBOX_DIR).exists() else []
    chosen = pick_consent_zip(zips) if zips else None
    if chosen is None:
        sys.exit("No consent ZIP found in the ingest folder. Pass a path explicitly.")
    return chosen


def main():
    ap = argparse.ArgumentParser(
        description="Flag bot/fraud/duplicate entries in a consent export ZIP.")
    ap.add_argument("zip", nargs="?", help="Consent export ZIP (default: newest in ingest).")
    ap.add_argument("-o", "--out", default="suspicious_report.csv",
                    help="Where to write the per-entry CSV report.")
    ap.add_argument("--no-blacklist", action="store_true",
                    help="Do not update the shared fraud blacklist (review only).")
    args = ap.parse_args()

    zip_path = Path(args.zip) if args.zip else _default_zip()
    print(f"Reading: {zip_path}\n")
    records = flag_zip(zip_path)
    print_summary(records)
    try:
        write_report(records, args.out)
        print(f"Wrote report -> {args.out}")
    except PermissionError:
        print(f"(Could not write {args.out} - file is open/locked. Close and re-run.)")

    if not args.no_blacklist:
        try:
            added = write_blacklist(records)
            print(f"Fraud blacklist -> {FRAUD_BLACKLIST_PATH} (+{added} new)")
        except PermissionError:
            print(f"(Could not write {FRAUD_BLACKLIST_PATH} - file is open/locked.)")


if __name__ == "__main__":
    main()
