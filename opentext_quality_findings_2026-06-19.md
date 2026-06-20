# Open-text quality screen: Q62 / Q63 (June 19, 2026 export)

Scope: the two free-text questions in the teen response survey.

- **Q62** "If you had to pick one for homework, which would you rather use? Your own AI account / Your school's AI account. Walk us through what would influence your decision."
- **Q63** same format, "An AI that keeps a record of your past work ... / An AI that keeps no record ...".

n = 418 responses (415 finished). Flagging was deliberately broad and teen-aware: informal tone, slang, lowercase, typos, and short-but-reasoned answers were **not** penalized. The target is junk and no-reasoning content, not casual teen writing.

Companion file: `data/opentext_quality_flags_2026-06-19.csv` (one row per flagged respondent, with the open-text reason, a confidence level, and the existing quality/fraud flags side by side).

## Headline

The existing quality filter never reads open-text content. Every signal it has (`speeding`, `recaptcha`, `attention`, `straightlining`, `incomplete`) is timing or metadata. So a respondent who paces normally, passes the attention checks, and submits "asdf" or pastes a ChatGPT answer is invisible to it.

Measured against this export:

- 220 of 418 respondents (52%) carry at least one open-text quality signal; 130 are high-confidence.
- Of those 220, only **22 (10%)** would be excluded by the current filter, and only **79 (36%)** trip any existing flag at all.
- Of the 130 high-confidence cases, only **11** would be excluded.

The open-text layer is the largest blind spot in the pipeline.

## What the low-quality responses look like

Ordered by how much they matter, not by raw count.

**1. Q63 answers that just restate the option (no reasoning): ~103 respondents.** This is the dominant problem. People typed the answer choice back instead of explaining. Many are lightly reworded copies of the option text, and the rewordings cluster suspiciously: 18 respondents independently wrote "al" instead of "ai" inside the same sentence ("an al that keeps records of my past..."), and 78 reworded the option from "your" to "my" (22 of them in the identical 13-word phrasing). Independent teens do not converge on the same misspelled sentence. This points to copying from a shared source or templated entry. Only 6 of the 103 are caught today.

**2. LLM-paste formatting: 9 respondents.** A tight cluster whose answers carry markers a teen typing in a box does not produce: indented sub-bullets ("Reliability of responses\n    Memory can improve relevance..."), em dashes, and a "Title then explanation" structure. Several share an identical Q62 ("School policies: Some schools require using their account for grading or tracking usage."). They cluster at reCAPTCHA 0.6 and durations of 198 to 314 seconds. The existing filter excludes exactly **1 of the 9**, because each has only the single reCAPTCHA flag and the rule needs two.

**3. Finished but both boxes blank: 18 respondents.** Completed the survey, left both required prose questions empty. Only 4 are caught today.

**4. Non-responsive one-word answers:** "No"/"No", "Nothing"/"Nothing", and similar across both questions. A handful, mostly uncaught.

**5. Gibberish / keyboard mash: 0.** None in this export. Worth keeping a cheap detector anyway.

Two signals I tested have poor precision and should stay review-only, never standalone exclusions: `off_topic` (a keyword test that misfires on real answers like "Either one. If I have given them permission ahead I would not mind.") and `ai_vocabulary` (flagging "furthermore"/"ultimately" catches articulate real teens). Both are in the CSV at low/medium confidence so you can judge them, but I would not wire them into an automated rule.

## Patterns in the existing heuristics

1. **No content signal exists.** The filter is structurally unable to catch content junk. This is the gap, not a tuning problem.

2. **reCAPTCHA is the one existing signal that co-travels with bad open-text.** Among flagged respondents, 98 of 220 (45%) errored on reCAPTCHA (`Q_RecaptchaStatus = error`); among clean respondents, 17 of 198 (9%). But the filter only reads the numeric `Q_RecaptchaScore` and ignores the `error` status entirely (115 responses errored in this export). The paste cluster sits at score 0.6, under the 0.7 cutoff but contributing only one flag.

3. **The two-flag combination rule is what lets the paste cluster through.** These respondents have exactly one behavioral flag (low reCAPTCHA). They need two to be excluded. A content flag would push them over.

4. **The consent fraud blacklist already knows 39 of these completers** (their cid or survey IP is on `data/fraud_blacklist.csv`, mostly fraud-farm IPs), but that list is only read at payment time. Their junk answers still flow into analysis because the quality/analysis stage never consults it.

## Proposed new heuristics

All of these feed `exclude_recommended` for analysis. They do not change the pay-everyone policy in `QUALITY_CONTROL.md`; payment stays decoupled.

Content layer (new, add to `quality_filter.py`):

- **`both_blank_finished`** finished with both required prose answers empty. Hard signal, standalone exclude. Catches the 18.
- **`llm_paste_format`** em or en dash, or an indented sub-bullet (`\n` followed by 2+ spaces then text) in a free-text box. Near-zero false positives for teens. Standalone exclude. Catches the 9; only 1 caught today.
- **`cross_respondent_duplicate`** same normalized open-text answer across different respondents, length-gated (>= ~7 words) so natural short collisions like "my own ai account" are ignored. Strong farm/paste signal. Catches the "al" and reworded-option clusters.
- **`option_echo` / `no_reasoning`** on "walk us through" prompts, the answer is only the choice or the piped option text (fuzzy match, tolerate "your"/"my" and "ai"/"al"). Frame as "unusable for the open-text analysis," not necessarily fraud. This is the ~103-respondent bucket.
- **`gibberish`** vowel ratio, repeated-char runs, keyboard rows. Cheap insurance even though it found nothing here.

Tuning the existing filter:

- **Use `Q_RecaptchaStatus = error` as a flag**, not just `score < 0.7`. It is the strongest existing correlate and is currently unused.
- **Promote the hard content signals** (`gibberish`, `cross_respondent_duplicate`, `both_blank_finished`, `llm_paste_format`) to standalone exclusion, bypassing the two-flag rule. These are near-certain junk; requiring a second flag is what hides the paste cluster.
- **Read `fraud_blacklist.csv` at the quality/analysis stage**, so the 39 known fraud-farm completers are `exclude_recommended` for analysis, not just suppressed at payment.

## Caveats

- "These are teens" cuts both ways: the screen is lenient on style, so a flag means low effort or non-human pattern, not bad writing. Review the CSV before excluding; the `off_topic` and `ai_vocabulary` rows in particular need eyes.
- The duplicate and option-echo counts could in principle include a Qualtrics piping artifact. The consistent rewordings and the shared "al" misspelling argue against pure piping, but confirm in the survey flow that the Q63 text box is not auto-populated from the selected choice.
- Counts are from the June 19 export only and will shift as responses come in.
