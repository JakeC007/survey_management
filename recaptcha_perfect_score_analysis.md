# Would a "perfect reCAPTCHA only" rule protect data quality?

**Question.** If we excluded every response without a perfect reCAPTCHA score (anything below 1.0), what would we lose, and is Jake right that many of the 0.8+ respondents are real people who shouldn't be dropped?

**Verdict.** The guess holds, with one correction. Respondents scoring 0.8–0.9 look like real, engaged people, and a perfect-score rule would discard 38 of them that the validated pipeline keeps as good data. But the deeper problem is that a perfect-score gate is the wrong instrument in *both* directions: it throws out good respondents *and* it waves through 91 flagged-junk responses that happen to score a perfect 1.0. reCAPTCHA score is a weak proxy for actual quality, which is why the existing filter uses it as one counting flag among many rather than a gate.

Source: canonical `quality_filter.py` run on `ingest/K12 Privacy and AI Extension_June 25, 2026_06.18.csv` (n = 642).

## What the perfect-score rule would actually do

A "drop everything below 1.0" rule removes **187 of 642 respondents (29%)**.

- **70 of those 187 (37%)** are clean on *every* independent quality signal: attention checks, speeding, straight-lining, gibberish, LLM-paste, duplicates, blacklist. No quality reason to drop them other than the score.
- **78 of those 187 (42%)** are responses the validated pipeline currently keeps as good data.
- Run the other direction: **91 respondents scored a perfect 1.0 yet are `exclude_recommended`** (20% of the perfect-score group). A perfect-score gate keeps all 91 of them.

So the rule over-excludes ~78 real respondents and under-excludes ~91 junk ones. It moves the sample sideways, not toward higher quality.

## reCAPTCHA score vs. independent quality signals

For each score band, "clean (other signals)" means zero independent flags. Median duration and attention-fail rate are independent tells of a real, engaged respondent.

| reCAPTCHA | n | Clean on other signals | Excluded by full rule | Attention fail | Speeding | Median duration |
|-----------|----|------------------------|-----------------------|----------------|----------|-----------------|
| 1.0 (perfect) | 455 | 70% | 20% | 4.6% | 3.7% | 532 s |
| 0.8–0.9 | 55 | 55% | 31% | 3.6% | 9.1% | 542 s |
| 0.7 | 11 | 27% | 73% | 9.1% | 9.1% | 346 s |
| < 0.7 | 121 | 31% | 69% | 2.5% | 20% | 241 s |

Per exact score:

| Score | n | Clean on other signals | Excluded by full rule | Median duration |
|-------|----|------------------------|-----------------------|-----------------|
| 1.0 | 455 | 320 (70%) | 91 | 532 s |
| 0.9 | 28 | 17 (61%) | 8 | 724 s |
| 0.8 | 27 | 13 (48%) | 9 | 439 s |
| 0.7 | 11 | 3 (27%) | 8 | 346 s |
| 0.6 | 58 | 16 (28%) | 42 | 219 s |
| 0.5 | 30 | 14 (47%) | 16 | 289 s |

## Reading the table

The 0.8–0.9 band behaves like the perfect-score group, not like the junk below it. Median time on task is **542 s vs. 532 s** for the perfect group, essentially identical, and the 0.9-only respondents actually spent *longer* (724 s median). Their attention-fail rate (3.6%) is lower than the perfect group's (4.6%). These are signatures of people who read the survey and answered it, not bots clicking through.

The real cliff is at **0.7 and below**. There, median duration collapses to ~240 s, speeding roughly quadruples to ~20%, and the share clean on other signals drops to ~30%. That is where reCAPTCHA score starts tracking actual junk, and it is why the pipeline already sets `recaptcha_min = 0.7`, counting (not hard-excluding) anything below it.

Where Jake's intuition needs the correction: "0.8+ are real" is too strong as a blanket claim. About **31% of the 0.8–0.9 band still trips two or more independent flags** and is excluded on its own merits. The point is not that every 0.8 respondent is real, it's that the reCAPTCHA score is not what tells you. The independent signals do the discriminating; among 0.8–0.9 respondents the pipeline keeps **38** and drops **17**, and a flat score rule cannot make that distinction.

## What a perfect-score rule costs and misses

| | Count |
|---|-------|
| Respondents dropped by perfect-score rule | 187 (29% of sample) |
| ...clean on every non-reCAPTCHA signal | 70 |
| ...currently kept as good data by validated pipeline | 78 |
| Good respondents kept by pipeline with imperfect score | 78 (incl. 38 in the 0.8–0.9 band) |
| Flagged-junk responses scoring a perfect 1.0 (rule would keep) | 91 |

## Recommendation

Keep treating reCAPTCHA as one counting signal, not a gate. The current configuration (`recaptcha_min = 0.7`, never a standalone hard exclude; `recaptcha_error` routes to Hold for human review) already captures the signal where it is informative and avoids the two failure modes a perfect-score rule creates. A perfect-score rule would cost ~78 real respondents, including ~38 in the 0.8–0.9 band, while leaving 91 flagged-junk responses in the sample. Net effect on data quality: negative.

If the goal is a defensible single number to report in a methods section, the honest framing is "responses below 0.7 were flagged and reviewed alongside other quality signals," not "only perfect-score responses were retained."

---

### Method notes

- Independent ("other") signals exclude both reCAPTCHA flags, so the cleanliness comparison is not circular: it asks whether 0.8–0.9 respondents look good *by evidence other than the score*.
- "Excluded by full rule" is the pipeline's `exclude_recommended` (≥ 2 independent flags, or all shown attention checks failed, or a hard signal like gibberish/duplicate/blacklist).
- Cross-respondent duplicate and blacklist signals were computed across the full June 25 export before the per-row pass, per the filter's design.
- Scores rounded to one decimal to absorb floating-point representation (Qualtrics reCAPTCHA v3 reports in 0.1 steps).
