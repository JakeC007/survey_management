# Quality-filter changes: open-text + fraud signals (June 19, 2026)

Four changes to the shared exclusion filter, all driven by the Q62/Q63 open-text
analysis. Implemented in the two vendored copies (`response_management/`,
`payment_management/`) and unit-tested. The canonical copy still needs the same
edit plus a sync (see "Canonical sync" at the end).

## What changed

1. **reCAPTCHA error is now a flag.** `Q_RecaptchaStatus == "error"` raises
   `flag_recaptcha_error`, separate from the numeric-score flag. It is the
   strongest correlate of junk open-text (errored on ~45% of flagged responses
   vs ~9% of clean) and was previously unused. It is an ordinary counting flag,
   not a hard exclusion: on its own it does not drop a row from analysis.

2. **Four content signals are now standalone (HARD) exclusions**, bypassing the
   two-flag rule: `gibberish`, `cross_respondent_duplicate`, `both_blank_finished`,
   `llm_paste_format`. Any one of them sets `exclude_recommended = True`.
   Requiring a second flag is exactly what let the LLM-paste cluster pass before.

3. **The fraud blacklist is read at the quality/analysis stage.** A completer
   whose cid or survey IP is on `data/fraud_blacklist.csv` now gets
   `flag_fraud_blacklist` + `exclude_recommended`, so the known fraud-farm
   completers drop from the analysis sample, not just from payment. Auto-discovers
   the blacklist near the module or export; opt-in (empty sets if no file found).

4. **`cross_respondent_duplicate` is built to never hold a real teen** (see QA).

New config keys live in `quality_config.json`: `recaptcha_error_is_flag`,
`open_text_cols`, `duplicate_min_words/chars`, `option_text`,
`option_echo_jaccard`, `fraud_blacklist_path`, `hard_exclude_signals`.

Architecture note: `cross_respondent_duplicate` and the blacklist are cross-row,
so `evaluate_export` and `from_dataframe` now do one corpus pre-pass before the
per-row pass. `from_dataframe` also coerces pandas NaN to "" so blank detection
works in the analysis repo.

## Effect on the June 19 export (418 responses, 382 completers)

Analysis: `exclude_recommended` = 125. New flag counts: recaptcha_error 115,
fraud_blacklist 60, cross_respondent_duplicate 28 respondents (30 rows, incl. a
duplicate completion), both_blank_finished 18, llm_paste_format 9, gibberish 0.

Payment buckets (HOLD_BUCKET_ENABLED=True, HOLD_ON_ANY_FLAG=True):
Pay/Hold/Fraud = **153 / 176 / 53** (was 309/20/53 before any flagging changes).
Fraud is unchanged: the blacklist still routes to Fraud (never pay); the new
signals route to Hold (human review before paying), per the IRB-covered decision.

**Heads-up on the Hold size.** 94 of the 176 holds are driven by
`recaptcha_error` alone, because `HOLD_ON_ANY_FLAG=True` holds on any single
flag. Many of those are likely real teens on privacy browsers or ad-blockers
(reCAPTCHA failing to load is not proof of a bot). They are held for review, not
denied, and `recaptcha_error` alone does not drop them from analysis. If 176 is
too many to review by hand, the clean options are: exempt `recaptcha_error` from
Hold-gating (keep it analysis/annotation-only), or set `HOLD_ON_ANY_FLAG=False`
so Hold tracks `exclude_recommended` instead (that gives Pay/Hold/Fraud ≈
266/63/53). Say the word and I'll wire either.

## cross_respondent_duplicate: QA (the "don't hold real teens" requirement)

The signal only fires when a respondent's answer is, after normalization,
**verbatim-identical to another respondent's**, is substantive (>= 7 words,
>= 30 chars), and is **not** an echo of an answer option (token Jaccard < 0.6
against `option_text`). Consequences, verified on the real export:

- **A unique answer can never be flagged.** Flagging requires a duplicate
  partner, so any teen who writes original words is safe by construction. Of
  the open-text answers, 262/278 (Q62) and 292/310 (Q63) are unique.
- **Option echoes are spared.** Excluding option restatements drops the flag
  set from 91 to 28 respondents: **63 real teens who simply retyped the answer
  option are not flagged.** (Unit-tested.)
- **All 28 flagged are genuine paste/farm**, hand-checked across 8 clusters:
  five are a farm paraphrasing the "keeps a record" option with the identical
  "Al"-for-"AI" misspelling, recurring on the 66.93.14.x subnet; two are
  ChatGPT-paste (indented sub-bullets / em dashes); one is an identical
  distinctive 15-word sentence. None look like independent teens.

Because the payment routing is Hold (review), not Fraud (auto-deny), even a
hypothetical false positive is seen by a human before any card decision.

Unit tests: `payment_management/test_quality_filter.py` (18 tests) covers
gibberish, llm-paste, both-blank, recaptcha-error, blacklist, and the
duplicate cases (option-echo NOT flagged, distinctive paste flagged, short
collision NOT flagged, single respondent NOT flagged, unique answer safe),
plus backward-compat (clean row passes; two behavioral flags still exclude).
All 18 pass; existing `test_fraud_payments.py` (7) and
`test_flag_suspicious_entries.py` (21) still pass.

## A real catch from reading the blacklist at the quality stage

cid `R_1YDG8v69eHzuRuz` completed twice, from two IPs; one (`47.148.17.88`) is
on the blacklist, the other is not. Payment maps one IP per cid and had the
clean one, so payment alone would not have caught it. The quality-stage
blacklist read flags the blacklisted-IP submission, drops it from analysis, and
lands the completer in Hold for review. This is the kind of duplicate-from-a-
farm-IP case the change is meant to surface.

## Canonical sync (required, I could not do it from here)

The canonical `quality_filter.py` and `quality_config.json` live in the parent
`AI Survey/` folder, which is not accessible from this workspace. I edited only
the two `survey_management` copies. To make the change stick and reach analysis:

1. Apply the same diff to `AI Survey/quality_filter.py` and
   `AI Survey/quality_config.json` (the two copies here are the reference).
2. Run `sync_qc.bat` so all three vendored copies (analysis repo included) match,
   then commit each repo. Otherwise the next sync reverts these edits.
3. **The analysis repo needs the blacklist to use signal #3.** Ensure
   `Teen AI Survey Analysis` has `data/fraud_blacklist.csv` available (copy it,
   or set `fraud_blacklist_path` in its config). Without it, `from_dataframe`
   runs fine but the blacklist flag never fires there.
4. `open_text_cols` and `option_text` are specific to this survey's Q62/Q63. If
   the instrument changes, update them in the config.
