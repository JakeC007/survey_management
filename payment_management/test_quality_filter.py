"""Unit tests for the open-text / fraud signals added to quality_filter.py.
Run: python -m unittest test_quality_filter -v"""
import unittest
import quality_filter as qf

CODES = ["ResponseId", "cid", "Finished", "Duration (in seconds)",
         "Q_RecaptchaScore", "Q_RecaptchaStatus", "IPAddress", "Q62", "Q63",
         "Q220", "Q218", "Q219"]
CFG = qf.load_thresholds()


def mkrow(**kw):
    r = {c: "" for c in CODES}
    r["Finished"] = "True"
    r["Duration (in seconds)"] = "600"
    r["Q_RecaptchaScore"] = "1.0"
    r.update(kw)
    return r


def ev(row, dup_ids=None, fraud_ids=None, fraud_ips=None):
    return qf.evaluate_row(row, CODES, CFG, dup_ids or set(),
                           fraud_ids or set(), fraud_ips or set())


class Gibberish(unittest.TestCase):
    def test_keyboard_mash_flagged(self):
        for junk in ["asdfasdf", "qwerty qwer", "zxcvzxcv", "lkjlkjlkj"]:
            self.assertTrue(qf.is_gibberish(junk), junk)

    def test_real_teen_text_not_gibberish(self):
        for ok in ["idk", "my own bc privacy", "I would use my own account",
                   "no record cuz im scared", "schools ai lol"]:
            self.assertFalse(qf.is_gibberish(ok), ok)


class LLMPaste(unittest.TestCase):
    def test_em_dash_flagged(self):
        self.assertTrue(qf.looks_llm_paste("Memory helps — but only if accurate."))

    def test_indented_bullet_flagged(self):
        self.assertTrue(qf.looks_llm_paste("Reliability of responses\n    Memory can improve relevance."))

    def test_plain_teen_text_not_flagged(self):
        self.assertFalse(qf.looks_llm_paste("my own account because i dont want school to see"))


class BothBlankFinished(unittest.TestCase):
    def test_both_blank_and_finished_excluded(self):
        r = ev(mkrow(Q62="", Q63="", Finished="True"))
        self.assertTrue(r["flag_both_blank"])
        self.assertTrue(r["exclude_recommended"])   # HARD signal, standalone

    def test_blank_but_not_finished_not_bothblank(self):
        r = ev(mkrow(Q62="", Q63="", Finished="False"))
        self.assertFalse(r["flag_both_blank"])


class RecaptchaError(unittest.TestCase):
    def test_error_status_flagged_and_counts(self):
        r = ev(mkrow(Q_RecaptchaStatus="error", Q62="my own account because privacy"))
        self.assertTrue(r["flag_recaptcha_error"])
        self.assertGreaterEqual(r["n_flags"], 1)

    def test_error_alone_does_not_exclude(self):
        # one soft flag only -> not analysis-excluded (needs >=2 or a HARD signal)
        r = ev(mkrow(Q_RecaptchaStatus="error", Q62="I would use my own account because of privacy",
                     Q63="no record because i dont trust it"))
        self.assertTrue(r["flag_recaptcha_error"])
        self.assertFalse(r["exclude_recommended"])


class FraudBlacklist(unittest.TestCase):
    def test_blacklisted_cid_excluded(self):
        r = ev(mkrow(cid="R_BAD", Q62="my own account because privacy"),
               fraud_ids={"R_BAD"})
        self.assertTrue(r["flag_fraud_blacklist"])
        self.assertTrue(r["exclude_recommended"])

    def test_blacklisted_ip_excluded(self):
        r = ev(mkrow(cid="R_OK", IPAddress="1.2.3.4", Q62="x y z reason here please"),
               fraud_ips={"1.2.3.4"})
        self.assertTrue(r["flag_fraud_blacklist"])


class CrossRespondentDuplicate(unittest.TestCase):
    OPT = "An AI that keeps a record of your past work to give faster, more personalized help"

    def _dups(self, rows):
        return qf.compute_duplicate_flags(rows, CODES, CFG)[0]

    def test_option_echo_not_flagged(self):
        # Two teens both restate the option verbatim -> NOT a duplicate (option echo).
        rows = [mkrow(cid="A", Q63=self.OPT), mkrow(cid="B", Q63=self.OPT)]
        self.assertEqual(self._dups(rows), set())

    def test_distinctive_paste_flagged(self):
        txt = "I would pick the personalized AI if I had many assignments in the same subject"
        rows = [mkrow(cid="A", Q63=txt), mkrow(cid="B", Q63=txt)]
        self.assertEqual(self._dups(rows), {"A", "B"})

    def test_short_collision_not_flagged(self):
        rows = [mkrow(cid="A", Q62="my own account"), mkrow(cid="B", Q62="my own account")]
        self.assertEqual(self._dups(rows), set())

    def test_single_respondent_not_flagged(self):
        txt = "I would pick the personalized AI if I had many assignments in the same subject"
        rows = [mkrow(cid="A", Q63=txt)]
        self.assertEqual(self._dups(rows), set())

    def test_unique_distinctive_answer_safe(self):
        # A real teen's unique, original long answer is never flagged.
        rows = [
            mkrow(cid="A", Q63="i would keep no record because i worry about my data being sold later on"),
            mkrow(cid="B", Q63="honestly the school one feels safer even if they can see my stuff sometimes"),
        ]
        self.assertEqual(self._dups(rows), set())


class BackwardCompat(unittest.TestCase):
    def test_clean_row_not_excluded(self):
        r = ev(mkrow(Q62="my own account because i value privacy a lot",
                     Q63="no record since i dont trust storing my data",
                     Q218="Definitely not okay"))
        self.assertFalse(r["exclude_recommended"])
        self.assertEqual(r["n_flags"], 0)

    def test_two_behavioral_flags_still_excluded(self):
        # speeding + recaptcha low = 2 flags -> excluded (pre-existing rule intact)
        r = ev(mkrow(**{"Duration (in seconds)": "100", "Q_RecaptchaScore": "0.2",
                        "Q62": "my own account because privacy",
                        "Q63": "no record because i dont trust it"}))
        self.assertTrue(r["exclude_recommended"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
