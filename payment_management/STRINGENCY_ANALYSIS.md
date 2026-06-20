# Fraud / Data-Quality Stringency Analysis

Export analyzed: `K12 Privacy and AI Extension_June 19, 2026_07.20` (407 raw teen-survey rows; 382 of them are matched completers in the payment ledger). Fraud blacklist: 276 response IDs, 78 IPs. Nothing has been marked paid yet, so tightening now costs no clawbacks.

## 1. What the current rules actually are

The pipeline runs two independent layers. They answer different questions and should not be collapsed.

### Layer A — data quality (`quality_filter.py` + `quality_config.json`)

Per-respondent flags, then a combination rule. A response is `exclude_recommended` when **2 or more** independent flags trip, or when every shown attention check failed.

| Flag | Threshold (current) | Hit rate on 407 |
|---|---|---|
| Speeding (total) | whole survey < **180 s** | part of 3.4% |
| Speeding (pages) | >= **2** pages each < **2.0 s** | part of 3.4% |
| Attention | >= **1** of the shown instructed-response checks failed (Q220/Q218/Q219) | 4.9% |
| Straight-lining | >= **1** matrix of >= 5 answered items all identical (3 comfort grids + attitudes grid) | 8.8% |
| reCAPTCHA | Qualtrics v3 score < **0.5** | 4.7% |
| Incomplete | not finished (informational; does not count toward exclusion) | 0.7% |
| Combination | exclude when **n_flags >= 2**, or all shown attention checks failed | **3.9% (16 of 407)** |

### Layer B — fraud / identity (`fraud_blacklist.csv`, built upstream by the consent screen)

A completer is routed to **Fraud** (never paid) if its consent `cid` **or** its survey IP is on the blacklist. Blacklist reasons, by count: 181 "N entries from one IP", 50 "IP shared by N emails", 30 "same email+child submitted more than once", 8 speed, 7 reCAPTCHA. This layer matches **13.3% (54 of 407)** of the raw export.

A third check sits in `manage_payments.py` itself: a delivery email on the throwaway-domain list routes to **Hold** (27 completers), with a `.edu`/`.k12`/`.sch.uk` whitelist so school addresses are never held.

### Current payment buckets (382 completers)

Pay **309**, Hold **20**, Fraud **53**. Pass rate ≈ 81% of completers; 84% of the raw export.

## 2. What our heuristics "think" vs Qualtrics' 72%

The comparison only works if you line up like with like, and the two systems are not measuring the same thing.

Qualtrics' "300 passed / 72%" is almost certainly its reCAPTCHA/bot score. The reCAPTCHA histogram is the tell: **295 of 407 responses (72.5%) scored exactly 1.0**, and the other 27.5% scored below. That 27.5% matches Qualtrics' ~28% rejection too closely to be accidental. Qualtrics is passing the perfect-score rows and failing the rest.

Our **quality filter** barely uses that signal. It only flags reCAPTCHA below 0.5 (4.7%) and otherwise scores *answer behavior*, so it excludes just 3.9%. That is not our filter disagreeing with Qualtrics about who is a bot. It is our filter looking at a different axis and our fraud layer (the consent-screen blacklist) doing the identity work instead.

Put them together and the current pipeline rejects **16.5% of the raw export (67 of 407)** vs Qualtrics' ~28%. So today we are the *more lenient* system overall, and the entire gap is the reCAPTCHA band between 0.5 and 1.0 that we currently ignore (76 responses sit in 0.5–0.9).

One coincidence to not over-read: making any single quality flag exclude (Scenario A below) lands the raw pass rate at 71%, almost exactly Qualtrics' 72%. Those are different response sets being cut, not agreement. Matching the headline number is not the same as matching the decision.

## 3. Stricter scenarios (reproducible via `simulate_stringency.py`)

Two knobs: quality thresholds, and whether a shared IP counts as fraud. `qExcl` = data-quality exclusions, `fraud` = identity/never-pay, REJECT = union.

| Scenario | qExcl | fraud | REJECT (raw 407) | PASS (raw) | REJECT (pay 382) |
|---|---|---|---|---|---|
| **Current** (on-disk) | 4% | 13% | 16% | **84%** | 18% |
| **A** any single flag excludes | 18% | 13% | 29% | **71%** | 31% |
| **B** A + reCAPTCHA<0.7, dur<300s, 1 fast page | 50% | 13% | 54% | **46%** | 57% |
| **C** B + any shared IP = fraud | 50% | 33% | 60% | **40%** | 64% |
| **D** reCAPTCHA<0.9, dur<300s, any-flag | 53% | 13% | 57% | **43%** | 60% |
| **E** D + shared IP = fraud (most extreme) | 53% | 33% | 63% | **37%** | 66% |

Single-lever sensitivity (quality exclusions only, fraud held at current blacklist):

| Change | Excludes |
|---|---|
| reCAPTCHA 0.5 → 0.7 | 5.9% |
| reCAPTCHA 0.5 → 0.9 | 7.6% |
| duration floor 180 → 300 s | 7.4% |
| duration floor 180 → 420 s | 10.8% |
| 2 fast pages → 1, page floor 2 → 3 s | 9.1% |
| **n_flags 2 → 1 (any flag)** | **18.9%** |

## 4. Reading the levers

**reCAPTCHA is the highest-value underused signal.** It is the axis Qualtrics rejects on and the one we discount. Raising the threshold to 0.7 is the cheapest move toward parity, and it belongs in the *fraud* logic, not just quality, since a low v3 score is a bot signal, not a careless-teen signal.

**IP clustering is real but the blunt version is dangerous.** 29.7% of responses share an IP with another response, and the worst farms are already on the blacklist (one IP had 14 entries). But a blanket "any shared IP = fraud" rule (Scenarios C/E, fraud jumps to 33%) will catch siblings in one household, a classroom behind one school NAT, and mobile-carrier CGNAT. For a K-12 teen sample that is a live false-positive risk. The consent screen's existing approach (flag an IP shared by 3+ *distinct emails*, or repeat email+child) is the safer threshold to tighten, not raw IP reuse.

**The aggressive quality scenarios (B–E) gut the sample.** Excluding 50–66% of an n≈400 study is only defensible if you have evidence of a specific mass-fraud event. Absent that, these throw away far more real teens than fraudsters.

## 5. Recommendation

Decide what you are protecting before turning a knob, because the two layers have different costs:

- **Protecting the money (fraud):** this is the layer worth being extreme on, and it is nearly free right now since nobody has been paid. Raise `recaptcha_min` to 0.7 and treat it as a fraud signal; tighten the consent-screen IP rule on *distinct emails per IP* rather than raw IP reuse. This catches the bot/farm band Qualtrics already rejects without nuking households.
- **Protecting the analysis (quality):** moving `exclude_min_flags` from 2 to 1 (any flag → Hold, ~18%) is a defensible, reviewable middle ground given the small n. Keep it as Hold/review, not auto-drop. Per `QUALITY_CONTROL.md` the default protocol is to pay every completer and treat exclusion as an analysis decision; the Hold bucket already deviates from that under IRB clearance, so confirm any payment-gating change is still covered.

Run `python simulate_stringency.py` after each new export to re-check these counts as the sample grows.

## 6. Implemented (June 19, 2026): moderate option, applied

The moderate tier was applied to four files. All changes are reversible via git.

Teen survey:

- `quality_filter.py` + `quality_config.json` (canonical in `AI Survey/`, synced to all three vendored copies via `sync_qc.bat`): `recaptcha_min` 0.5 → 0.7.
- `payment_management/manage_payments.py`: `HOLD_ON_ANY_FLAG` False → True. This routes any flagged completer (n_flags >= 1) to Hold for review. It is set here, not as `exclude_min_flags` in the shared config, so the analysis repo's drop rule stays at 2 flags.

Effect on 382 completers (08.01 export): Pay/Hold/Fraud moves 309/20/53 → **240/89/53**. 69 move Pay→Hold; fraud unchanged. Nobody is paid yet, so no clawbacks.

Consent screen (`consent_management/flag_suspicious_entries.py`):

- `recaptcha_min` 0.5 → 0.7 (soft bot flag).
- `fast_seconds` 60 → 76 (soft speeding floor; see below).
- new `fast_hard_seconds` = 30 (hard tripwire that blacklists alone).

Effect on 992 finished consents: suspicious entries 276 → ~345 (+45 from reCAPTCHA, +24 from the wider soft-speeding band; the 30 s hard tripwire fires 0 today because the fastest real completion is 46 s).

### Why 76 s, and why SD-based floors do not work raw

Consent finish time is severely right-skewed (median 145 s, but mean 360 s, SD 3016 s, CV 8.37) because of entries left open for hours (max 81,151 s ≈ 22.5 h). So mean − k·SD goes negative and flags nobody:

| Method | n | mean (s) | SD (s) | mean − 1 SD | mean − 2 SD |
|---|---|---|---|---|---|
| full (no exclusion) | 992 | 360.3 | 3016.2 | −2655.9 | −5672.1 |
| Tukey IQR 1.5× trim (drop > 422 s) | 916 | 157.9 | 82.5 | **75.4** | −7.1 |
| trim top 10% | 892 | 151.5 | 73.5 | **78.0** | **4.5** |
| trim top 20% | 793 | 133.0 | 53.6 | **79.4** | **25.8** |

Log-normal fit (correct model for completion times; floors positive by construction): exp(μ − 1σ) = **76 s**, exp(μ − 2σ) = **36 s**.

The IQR-trim −1 SD (75 s) and the log-normal −1σ (76 s) agree, so the soft floor is set to **76 s**. The hard tripwire (30 s) sits below the log-normal −2σ (36 s) and below the 46 s observed minimum, so it blacklists nobody now but trips on genuine bot speed.

New consent exclusions at each candidate floor (HARD = speed blacklists alone; SOFT = needs a second signal):

| Floor | T (s) | under T | new if HARD | new if SOFT |
|---|---|---|---|---|
| log-normal − 2σ | 36 | 0 | +0 | +0 |
| IQR-trim mean − 1 SD | 75 | 118 | +118 | +0 to +27 |
| log-normal − 1σ (chosen soft) | 76 | 124 | +124 | +0 to +27 |
| 90 s (prior draft) | 90 | 208 | +208 | +27 |
| 30 s (chosen hard) | 30 | 0 | +0 | +0 |

Hard-flagging on speed alone is rejected for consent: at 76 s it would blacklist 124 people whose only signal is finishing in under 76 s, almost all real parents. The soft floor adds only corroborated exclusions.

Note: the canonical `quality_filter.py`/`quality_config.json` live in `AI Survey/` and are pushed to the vendored copies by `sync_qc.bat`. Edit the canonical files, re-run the sync, and commit each repo separately, or vendored edits revert.
