"""
Tests for flag_suspicious_entries.py

Run from the consent_management folder:
    python -m unittest test_flag_suspicious_entries -v

The synthetic fixtures mimic the human-readable label header that
load_csv_from_zip produces, so flag_entries sees exactly the column names it
would in production. The final test is a smoke test against the real June 16
consent export if it is present in the ingest folder.
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

import flag_suspicious_entries as fse
from send_survey_emails import INBOX_DIR, pick_consent_zip


# Column labels as they appear in a real Qualtrics consent export (label row).
COLS = {
    "rid": "Response ID",
    "recorded": "Recorded Date",
    "finished": "Finished",
    "parent": "Parent/Guardian Full Name",
    "child": "Child Participant's Full Name",
    "pconsent": "I confirm that I have read the study information sheet",
    "psig": "Typed signature - Parent",
    "assent": "I understand what the researchers asked me to do",
    "csig": "Typed signature - Child",
    "email": "What email address should we send the survey link to?",
    "ip": "IP Address",
    "recaptcha": "Q_RecaptchaScore",
    "duration": "Duration (in seconds)",
}


def make_df(rows: list[dict]) -> pd.DataFrame:
    """Build a label-headed DataFrame from compact row dicts, filling sane
    defaults so a 'clean' row trips no flags."""
    records = []
    for i, r in enumerate(rows):
        records.append({
            COLS["rid"]: r.get("rid", f"R{i:03d}"),
            COLS["recorded"]: r.get("recorded", "2026-06-16 10:00:00"),
            COLS["finished"]: r.get("finished", "True"),
            COLS["parent"]: r.get("parent", f"Parent{i}"),
            COLS["child"]: r.get("child", f"Child{i}"),
            COLS["pconsent"]: r.get("pconsent", "Yes"),
            COLS["psig"]: r.get("psig", f"Parent{i}"),
            COLS["assent"]: r.get("assent", "Yes"),
            COLS["csig"]: r.get("csig", f"Child{i}"),
            COLS["email"]: r.get("email", f"person{i}@gmail.com"),
            COLS["ip"]: r.get("ip", f"10.0.0.{i}"),
            COLS["recaptcha"]: r.get("recaptcha", "0.9"),
            COLS["duration"]: r.get("duration", "300"),
        })
    return pd.DataFrame(records).astype(str)


def by_id(records):
    return {r["response_id"]: r for r in records}


class TestCleanData(unittest.TestCase):
    def test_clean_rows_not_flagged(self):
        df = make_df([{}, {}, {}])
        recs = fse.flag_entries(df)
        self.assertEqual(len(recs), 3)
        self.assertTrue(all(not r["suspicious"] for r in recs))

    def test_incomplete_rows_dropped(self):
        df = make_df([{"rid": "A"}, {"rid": "B", "finished": "False"}])
        recs = fse.flag_entries(df)
        self.assertEqual([r["response_id"] for r in recs], ["A"])


class TestHardFlags(unittest.TestCase):
    def test_ip_farm(self):
        # 5 finished entries from one IP, distinct emails/kids -> farm.
        rows = [{"rid": f"F{i}", "ip": "203.0.113.7",
                 "email": f"a{i}@gmail.com", "child": f"Kid{i}"} for i in range(5)]
        recs = by_id(fse.flag_entries(make_df(rows)))
        self.assertTrue(all(recs[f"F{i}"]["flag_ip_farm"] for i in range(5)))
        self.assertTrue(all(recs[f"F{i}"]["suspicious"] for i in range(5)))

    def test_ip_farm_just_below_threshold_not_flagged(self):
        # 4 siblings on one home IP: shared but not a farm, no other signal.
        rows = [{"rid": f"S{i}", "ip": "198.51.100.9",
                 "email": "mom@gmail.com" if i == 0 else f"s{i}@gmail.com",
                 "child": f"Sib{i}"} for i in range(4)]
        recs = by_id(fse.flag_entries(make_df(rows)))
        for i in range(4):
            self.assertFalse(recs[f"S{i}"]["flag_ip_farm"])

    def test_duplicate_same_child_flagged(self):
        rows = [
            {"rid": "D1", "email": "dup@gmail.com", "child": "Sam Lee", "ip": "1.1.1.1"},
            {"rid": "D2", "email": "dup@gmail.com", "child": "Sam Lee", "ip": "2.2.2.2"},
        ]
        recs = by_id(fse.flag_entries(make_df(rows)))
        self.assertTrue(recs["D1"]["flag_dup_submission"])
        self.assertTrue(recs["D2"]["flag_dup_submission"])
        self.assertTrue(recs["D1"]["suspicious"])

    def test_multichild_parent_not_flagged(self):
        # Same email, DIFFERENT children = legitimate siblings.
        rows = [
            {"rid": "M1", "email": "mom@gmail.com", "child": "Ava Reed", "ip": "3.3.3.1"},
            {"rid": "M2", "email": "mom@gmail.com", "child": "Ben Reed", "ip": "3.3.3.2"},
        ]
        recs = by_id(fse.flag_entries(make_df(rows)))
        self.assertFalse(recs["M1"]["flag_dup_submission"])
        self.assertFalse(recs["M1"]["suspicious"])


class TestSoftFlags(unittest.TestCase):
    def test_single_soft_flag_not_suspicious(self):
        # Low reCAPTCHA alone is not enough.
        recs = by_id(fse.flag_entries(make_df([{"rid": "L1", "recaptcha": "0.2"}])))
        self.assertTrue(recs["L1"]["flag_low_recaptcha"])
        self.assertFalse(recs["L1"]["suspicious"])

    def test_two_soft_flags_suspicious(self):
        # Low reCAPTCHA + too fast = two soft flags -> suspicious.
        recs = by_id(fse.flag_entries(
            make_df([{"rid": "L2", "recaptcha": "0.2", "duration": "30"}])))
        self.assertTrue(recs["L2"]["suspicious"])
        self.assertEqual(recs["L2"]["n_soft_flags"], 2)

    def test_same_signature_soft_flag(self):
        recs = by_id(fse.flag_entries(
            make_df([{"rid": "G1", "psig": "Same Name", "csig": "Same Name"}])))
        self.assertTrue(recs["G1"]["flag_same_signature"])

    def test_signature_minor_difference_not_flagged(self):
        # Middle initial / spacing differences must NOT trip same_signature.
        recs = by_id(fse.flag_entries(
            make_df([{"rid": "G2", "parent": "Jane A Doe", "psig": "Jane A. Doe",
                      "child": "Jane Doe", "csig": "Jane  Doe"}])))
        self.assertFalse(recs["G2"]["flag_same_signature"])


class TestNameSwap(unittest.TestCase):
    # parent first name reused as the child's surname (token swap)
    SWAP = {"parent": "Gamey Locale", "child": "Doole Gamey"}

    def test_detected_but_not_suspicious_alone(self):
        recs = by_id(fse.flag_entries(make_df([dict(rid="N1", **self.SWAP)])))
        self.assertTrue(recs["N1"]["flag_name_swap"])
        self.assertFalse(recs["N1"]["suspicious"])   # contextual flag never fires alone

    def test_escalates_with_a_technical_signal(self):
        recs = by_id(fse.flag_entries(
            make_df([dict(rid="N2", recaptcha="0.2", **self.SWAP)])))
        self.assertTrue(recs["N2"]["flag_name_swap"])
        self.assertTrue(recs["N2"]["suspicious"])    # name_swap + low reCAPTCHA

    def test_name_swap_plus_signature_only_not_suspicious(self):
        # Two contextual flags, no technical signal -> review only, not suspicious.
        recs = by_id(fse.flag_entries(make_df([
            dict(rid="N3", psig="Same One", csig="Same One", **self.SWAP)])))
        self.assertTrue(recs["N3"]["flag_name_swap"])
        self.assertTrue(recs["N3"]["flag_same_signature"])
        self.assertFalse(recs["N3"]["suspicious"])

    def test_normal_shared_surname_is_not_swap(self):
        recs = by_id(fse.flag_entries(make_df([
            {"rid": "N4", "parent": "Iris Fels", "child": "Sabelle Fels"}])))
        self.assertFalse(recs["N4"]["flag_name_swap"])

    def test_patronymic_alone_is_safe(self):
        # South-Indian-style: child carries the father's given name. On its own
        # (no technical signal) this must never be marked suspicious.
        recs = by_id(fse.flag_entries(make_df([
            {"rid": "N5", "parent": "Suresh Kumar", "child": "Anjali Suresh"}])))
        self.assertTrue(recs["N5"]["flag_name_swap"])
        self.assertFalse(recs["N5"]["suspicious"])

    def test_single_token_names_not_swap(self):
        recs = by_id(fse.flag_entries(make_df([
            {"rid": "N6", "parent": "Madonna", "child": "Doole Gamey"}])))
        self.assertFalse(recs["N6"]["flag_name_swap"])


class TestThresholdOverride(unittest.TestCase):
    def test_override_lowers_farm_threshold(self):
        rows = [{"rid": f"T{i}", "ip": "9.9.9.9",
                 "email": f"t{i}@gmail.com", "child": f"K{i}"} for i in range(3)]
        recs = by_id(fse.flag_entries(make_df(rows), {"ip_farm_min": 3}))
        self.assertTrue(all(recs[f"T{i}"]["flag_ip_farm"] for i in range(3)))


class TestRecordHasChildName(unittest.TestCase):
    def test_child_name_in_record(self):
        recs = by_id(fse.flag_entries(make_df([{"rid": "C1", "child": "Ada Byron"}])))
        self.assertEqual(recs["C1"]["child_name"], "Ada Byron")


class TestBlacklist(unittest.TestCase):
    def _tmp(self):
        d = tempfile.mkdtemp()
        return Path(d) / "fraud_blacklist.csv"

    def test_write_only_suspicious_and_roundtrip(self):
        path = self._tmp()
        recs = [
            {"response_id": "R1", "ip": "1.1.1.1", "child_name": "A",
             "delivery_email": "a@x.com", "reasons": "ip_farm", "suspicious": True},
            {"response_id": "R2", "ip": "2.2.2.2", "child_name": "B",
             "delivery_email": "b@x.com", "reasons": "", "suspicious": False},
        ]
        added = fse.write_blacklist(recs, path)
        self.assertEqual(added, 1)
        ids, ips = fse.load_blacklist_sets(path)
        self.assertEqual(ids, {"R1"})
        self.assertEqual(ips, {"1.1.1.1"})

    def test_upsert_preserves_manual_rows_and_dedupes(self):
        path = self._tmp()
        # Seed a hand-added manual row.
        import csv
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=fse.BLACKLIST_COLS)
            w.writeheader()
            w.writerow({"response_id": "MANUAL1", "ip": "9.9.9.9",
                        "child_name": "", "delivery_email": "", "reasons": "by hand",
                        "source": "manual", "added_at": "2026-06-16T00:00:00"})
        rec = {"response_id": "R1", "ip": "1.1.1.1", "child_name": "A",
               "delivery_email": "a@x.com", "reasons": "ip_farm", "suspicious": True}
        fse.write_blacklist([rec], path)
        # Writing R1 again must not duplicate it; manual row must survive.
        added = fse.write_blacklist([rec], path)
        self.assertEqual(added, 0)
        rows = fse.load_blacklist(path)
        ids = [r["response_id"] for r in rows]
        self.assertIn("MANUAL1", ids)
        self.assertEqual(ids.count("R1"), 1)
        manual = next(r for r in rows if r["response_id"] == "MANUAL1")
        self.assertEqual(manual["source"], "manual")


class TestRealExportSmoke(unittest.TestCase):
    def test_real_consent_zip_if_present(self):
        zips = sorted(Path(INBOX_DIR).glob("*.zip")) if Path(INBOX_DIR).exists() else []
        chosen = pick_consent_zip(zips) if zips else None
        if chosen is None:
            self.skipTest("No consent ZIP in ingest folder.")
        recs = fse.flag_zip(chosen)
        self.assertGreater(len(recs), 0)
        # Sanity: some but not all should be flagged on the real pilot data.
        flagged = [r for r in recs if r["suspicious"]]
        self.assertGreater(len(flagged), 0)
        self.assertLess(len(flagged), len(recs))


if __name__ == "__main__":
    unittest.main(verbosity=2)
