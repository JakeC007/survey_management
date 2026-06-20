"""
Tests for the Fraud bucket in manage_payments.py

Run from the payment_management folder:
    python -m unittest test_fraud_payments -v

These cover the new fraud logic without needing Outlook or a real tracker:
fraud routing by cid and by IP, the bucket split (fraud takes priority over
Pay and Hold), and reading the shared blacklist CSV.
"""

import csv
import tempfile
import unittest
from pathlib import Path

import pandas as pd

import manage_payments as mp


def tracker(rows):
    """Minimal Participants frame: one dict per row, defaults filled."""
    recs = []
    for i, r in enumerate(rows):
        recs.append({
            "response_id": r.get("cid", f"R{i}"),
            "child_name": r.get("child_name", f"Kid{i} Last"),
            "delivery_email": r.get("email", f"k{i}@gmail.com"),
            # After cutoff so the row is payable.
            "survey_completed_at": r.get("completed_at", "2026-06-15 10:00:00"),
        })
    return pd.DataFrame(recs).astype(str)


def by_cid(rows):
    return {r["cid"]: r for r in rows}


class TestFraudRouting(unittest.TestCase):
    def test_cid_blacklisted_is_fraud(self):
        df = tracker([{"cid": "BAD"}, {"cid": "GOOD"}])
        rows = by_cid(mp.build_payment_rows(
            df, {}, fraud_ids={"BAD"}, fraud_ips=set(),
            fraud_reasons={"BAD": "ip_farm"}, ip_by_cid={}))
        self.assertEqual(rows["BAD"]["fraud"], "yes")
        self.assertIn("cid blacklisted", rows["BAD"]["fraud_reason"])
        self.assertEqual(rows["GOOD"]["fraud"], "no")

    def test_blacklisted_ip_is_fraud_even_if_cid_clean(self):
        df = tracker([{"cid": "C1"}])
        rows = by_cid(mp.build_payment_rows(
            df, {}, fraud_ids=set(), fraud_ips={"66.66.66.66"},
            fraud_reasons={}, ip_by_cid={"C1": "66.66.66.66"}))
        self.assertEqual(rows["C1"]["fraud"], "yes")
        self.assertIn("survey IP", rows["C1"]["fraud_reason"])

    def test_clean_ip_and_cid_not_fraud(self):
        df = tracker([{"cid": "C1"}])
        rows = by_cid(mp.build_payment_rows(
            df, {}, fraud_ids={"OTHER"}, fraud_ips={"1.2.3.4"},
            fraud_reasons={}, ip_by_cid={"C1": "5.6.7.8"}))
        self.assertEqual(rows["C1"]["fraud"], "no")


class TestBucketPriority(unittest.TestCase):
    def test_fraud_excluded_from_pay_and_hold(self):
        # A fraud row that ALSO trips the quality exclusion must land only in
        # Fraud, never in Hold or Pay.
        rows = [
            {"cid": "F1", "fraud": "yes", "fraud_reason": "cid blacklisted",
             "exclude_recommended": "yes", "paid": ""},
            {"cid": "H1", "fraud": "no", "exclude_recommended": "yes", "paid": ""},
            {"cid": "P1", "fraud": "no", "exclude_recommended": "no", "paid": ""},
        ]
        pay, hold, fraud = mp.bucket_rows(rows)
        self.assertEqual([r["cid"] for r in fraud], ["F1"])
        self.assertEqual([r["cid"] for r in hold], ["H1"])
        self.assertEqual([r["cid"] for r in pay], ["P1"])

    def test_paid_rows_drop_out_of_all_buckets(self):
        rows = [{"cid": "X", "fraud": "no", "exclude_recommended": "no", "paid": "yes"}]
        pay, hold, fraud = mp.bucket_rows(rows)
        self.assertEqual((pay, hold, fraud), ([], [], []))


class TestLoadBlacklist(unittest.TestCase):
    def test_reads_ids_ips_and_reasons(self):
        path = Path(tempfile.mkdtemp()) / "fraud_blacklist.csv"
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=["response_id", "ip", "child_name",
                                               "delivery_email", "reasons",
                                               "source", "added_at"])
            w.writeheader()
            w.writerow({"response_id": "R1", "ip": "1.1.1.1", "reasons": "ip_farm",
                        "source": "auto", "added_at": "x"})
        ids, ips, reasons = mp.load_fraud_blacklist(path)
        self.assertEqual(ids, {"R1"})
        self.assertEqual(ips, {"1.1.1.1"})
        self.assertEqual(reasons["R1"], "ip_farm")

    def test_missing_file_is_empty(self):
        ids, ips, reasons = mp.load_fraud_blacklist(Path("/no/such/file.csv"))
        self.assertEqual((ids, ips, reasons), (set(), set(), {}))


if __name__ == "__main__":
    unittest.main(verbosity=2)
