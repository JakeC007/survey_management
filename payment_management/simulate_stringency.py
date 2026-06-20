#!/usr/bin/env python3
"""
simulate_stringency.py — what-if simulator for the Teen AI Survey quality/fraud rules.

Sits beside manage_payments.py and reuses its exact logic:
  - quality_filter.py for the data-quality flags + exclude_recommended rule
  - manage_payments.load_fraud_blacklist() for the identity/fraud blacklist

It does NOT change any thresholds on disk. It re-evaluates the response export
under several candidate threshold sets and prints how many responses each would
PASS / REJECT, on two populations:

  RAW   = every row in the teen-survey export (apples-to-apples with Qualtrics'
          own "X% passed their quality check" number).
  PAY   = the completer population the payment pipeline actually acts on
          (rows whose cid is present in the export AND appear as a completed,
          post-cutoff participant in the payment ledger). This is where money
          is at stake.

Two independent levers are simulated:
  * QUALITY thresholds (passed to quality_filter via overrides).
  * A FRAUD rule:
      'blacklist'           — current behaviour: cid OR IP on fraud_blacklist.csv
      'blacklist+sharedip'  — ALSO treat any response whose IP is shared by >1
                              response in this export as fraud (stricter).

Usage:
    python simulate_stringency.py                 # newest export in ../ingest
    python simulate_stringency.py --export PATH    # specific zip/csv
"""
from __future__ import annotations

import argparse
import io
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import quality_filter as qf            # noqa: E402
import manage_payments as mp           # noqa: E402

N_BASELINE_NOTE = "Qualtrics benchmark: ~300 passed / ~72% (so ~28% rejected)."


def _export_to_csv(path: Path) -> Path:
    """Return a CSV path for a .zip or .csv export (extracting the zip to temp)."""
    if path.suffix.lower() != ".zip":
        return path
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.lower().endswith(".csv")]
        if not names:
            sys.exit(f"No CSV inside {path}")
        tmp = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb")
        tmp.write(z.read(names[0]))
        tmp.close()
        return Path(tmp.name)


def _key(row) -> str:
    return str(row.get("cid", "")).strip() or str(row.get("ResponseId", "")).strip()


# Candidate threshold sets. Each is (label, quality-overrides, fraud-rule).
# overrides=None means "use the on-disk config" (the current rule).
SCENARIOS = [
    ("CURRENT (on-disk config)", None, "blacklist"),
    ("A: any single flag excludes",
     {"exclude_min_flags": 1}, "blacklist"),
    ("B: A + recaptcha<0.7, dur<300, page<3x1",
     {"exclude_min_flags": 1, "recaptcha_min": 0.7,
      "total_duration_floor_sec": 300, "page_floor_sec": 3.0,
      "min_fast_pages_to_flag": 1}, "blacklist"),
    ("C: B + shared-IP counted as fraud",
     {"exclude_min_flags": 1, "recaptcha_min": 0.7,
      "total_duration_floor_sec": 300, "page_floor_sec": 3.0,
      "min_fast_pages_to_flag": 1}, "blacklist+sharedip"),
    ("D: recaptcha<0.9, dur<300, any-flag",
     {"exclude_min_flags": 1, "recaptcha_min": 0.9,
      "total_duration_floor_sec": 300, "page_floor_sec": 3.0,
      "min_fast_pages_to_flag": 1}, "blacklist"),
    ("E: D + shared-IP fraud (most stringent)",
     {"exclude_min_flags": 1, "recaptcha_min": 0.9,
      "total_duration_floor_sec": 300, "page_floor_sec": 3.0,
      "min_fast_pages_to_flag": 1}, "blacklist+sharedip"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", help="Response export (zip or csv). Default: newest in ../ingest.")
    args = ap.parse_args()

    export = mp.find_response_export(args.export)
    if export is None:
        sys.exit("No response export found. Pass --export PATH.")
    csv_path = _export_to_csv(Path(export))

    codes, rows = qf.load_qualtrics_csv(str(csv_path))
    n_raw = len(rows)

    # Fraud blacklist + shared-IP set (within this export).
    fraud_ids, fraud_ips, _reasons = mp.load_fraud_blacklist()
    ipc = Counter(str(r.get("IPAddress", "")).strip() for r in rows
                  if str(r.get("IPAddress", "")).strip())
    shared_ips = {ip for ip, c in ipc.items() if c > 1}

    # PAY population = export rows whose cid is a completer in the ledger.
    pay_keys = set()
    if mp.PAYMENT_TRACKER_PATH.exists():
        import pandas as pd
        led = pd.read_excel(mp.PAYMENT_TRACKER_PATH, sheet_name="Payments",
                            dtype=str).fillna("")
        ledger_cids = {str(c).strip() for c in led["cid"]}
        for r in rows:
            cid = str(r.get("cid", "")).strip()
            if cid and cid in ledger_cids:
                pay_keys.add(_key(r))
    n_pay = len(pay_keys) if pay_keys else 0

    def evaluate(overrides, fraud_rule):
        recs = qf.evaluate_export(str(csv_path), overrides=overrides)[0]
        qexcl = {r["id"] for r in recs if r["exclude_recommended"]}
        fraud = set()
        for r in rows:
            cid = str(r.get("cid", "")).strip()
            ip = str(r.get("IPAddress", "")).strip()
            bl = (cid in fraud_ids) or (ip in fraud_ips)
            sh = ip in shared_ips
            if fraud_rule == "blacklist" and bl:
                fraud.add(_key(r))
            elif fraud_rule == "blacklist+sharedip" and (bl or sh):
                fraud.add(_key(r))
        reject_all = qexcl | fraud
        reject_pay = (reject_all & pay_keys) if pay_keys else set()
        return qexcl, fraud, reject_all, reject_pay

    print(f"\nExport      : {Path(export).name}")
    print(f"Raw rows    : {n_raw}    Completer (PAY) rows in export: {n_pay}")
    print(f"Fraud lists : {len(fraud_ids)} ids, {len(fraud_ips)} ips, "
          f"{len(shared_ips)} shared-IPs in export")
    print(N_BASELINE_NOTE)
    print("=" * 100)
    hdr = (f"{'scenario':42s}{'qExcl':>11}{'fraud':>11}"
           f"{'REJECT(raw)':>14}{'PASS(raw)':>13}{'REJECT(pay)':>13}")
    print(hdr)
    print("-" * 100)
    for label, ov, fr in SCENARIOS:
        qexcl, fraud, rej_raw, rej_pay = evaluate(ov, fr)
        pass_raw = n_raw - len(rej_raw)
        line = (f"{label:42s}"
                f"{len(qexcl):>6} {100*len(qexcl)/n_raw:>4.0f}%"
                f"{len(fraud):>6} {100*len(fraud)/n_raw:>4.0f}%"
                f"{len(rej_raw):>7} {100*len(rej_raw)/n_raw:>4.0f}%"
                f"{pass_raw:>6} {100*pass_raw/n_raw:>4.0f}%")
        if n_pay:
            line += f"{len(rej_pay):>6} {100*len(rej_pay)/n_pay:>4.0f}%"
        print(line)
    print("-" * 100)
    print("qExcl = data-quality exclusions; fraud = identity/never-pay; "
          "REJECT = union of the two.\n")


if __name__ == "__main__":
    main()
