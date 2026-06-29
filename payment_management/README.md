# payment_management

Owns everything about **who gets paid** for the Teen AI Survey. This is the
third folder of the survey-management system:

- `consent_management/` sends invitations and builds the master tracker.
- `response_management/` matches completions (by `cid`) and sends reminders.
- `payment_management/` (this folder) decides and tracks gift-card payment.

## What `manage_payments.py` does

1. Reads the **shared master tracker** (`../data/participant_tracker_auto.xlsx`),
   the same data source the other two folders use. A participant is payable when
   they have **both consent and a completed survey**. Both already live in the
   tracker: each row's `response_id` is the consent Response ID (= `cid`), and
   `survey_completed_at` is stamped by `manage_responses.py` only when a
   completed response's `cid` matches that row. So the "consent AND complete,
   merged on `cid`" check is just: a tracker row with `survey_completed_at` set.
2. Runs the shared **`quality_filter.py`** on the response-survey export (the
   Qualtrics teen-survey ZIP or CSV in `../ingest/`). The filter keys each
   respondent by `cid`, so its flags merge straight onto the tracker rows.
3. Upserts the master ledger **`../data/payment_tracker.xlsx`** (sheet
   `Payments`). One row per completed participant, with the quality flags and a
   hand-editable `paid` column. **Re-runs never overwrite your manual edits** to
   `paid`, `paid_date`, `paid_amount`, or `notes` (merged back by `cid`).
4. Writes **`../data/payment_report_unpaid.xlsx`** — everyone completed but not
   yet marked paid, split into three sheets:
   - **`Pay (no hold)`** — no exclusion recommended, not on the fraud blacklist.
   - **`Hold (flagged)`** — any quality flag from the filter (`n_flags >= 1`,
     the current `HOLD_ON_ANY_FLAG` setting; see Policy switches), **or** a
     delivery email on a known throwaway/disposable domain.
   - **`Fraud (do not pay)`** — `cid` or survey IP on the fraud blacklist.

   Each row carries the first name and the delivery email the survey was sent to.

## Fraud bucket vs Hold bucket

These are two different judgements and are decided independently:

- **Hold** is a **data-quality** call on a *real* participant's answers
  (speeding, failed attention checks, straight-lining), from `quality_filter.py`.
  Held teens can still be paid after review; exclusion is an analysis decision.
  A delivery email on a known throwaway/disposable domain
  (`../data/email_throwaway_domains.txt`, minus the whitelist) also lands here:
  we never auto-send a card to a disposable address, but it gets a human review
  rather than an outright fraud verdict, so a real teen on an odd domain isn't
  silently denied. Email typos (e.g. `gmil.com`) are corrected upstream, and
  institution domains (`.edu`, schools, exact entries like `lsoc.org`) are
  whitelisted, so neither lands in Hold for this reason.
- **Fraud** is an **identity** call made upstream at consent time by
  `consent_management/`'s suspicious-entry screen (bot / fraud-farm / duplicate
  signals). Those response IDs and the IPs they came from are written to
  **`../data/fraud_blacklist.csv`**. At payment time a completer is routed to
  the Fraud sheet — and **never** paid — if **either** its `cid` **or** the IP
  the survey was taken from is on that list. This catches a scammer who
  completes the survey after somehow getting hold of a link.

Fraud takes priority: a blacklisted row never appears in Pay or Hold, even if it
would otherwise be clean. If a fraud-matched person is *already* marked paid, the
run prints a loud warning so you can stop or claw back the card.

### Reviewer "keep" override (`../data/review_state.csv`)

The review queue in `examine_indv/` lets a human step through every flagged
completer and either mark them fraud or **keep** them. A keep is recorded in
**`../data/review_state.csv`** (`cid, decision=cleared, …`). At startup
`manage_payments.py` loads those cids (`load_kept_cids`) and `is_held()` forces
them out of Hold into Pay — so a kept completer stays payable no matter how many
times the build re-runs and re-computes quality flags. This is the Hold-side
mirror of the fraud blacklist: the blacklist forces cids *into* Fraud, the keep
list forces cids *out of* Hold. It never overrides Fraud (fraud is decided first)
and never moves anyone *into* Hold. The file is plain CSV and hand-editable;
remove a row (or add one with a non-`cleared` decision) to send a person back to
normal Hold/Pay evaluation. `examine_indv` also moves the row Hold → Pay in
`payment_report_unpaid.xlsx` at clear time, so the dashboard reflects it without
waiting for the next build.

### The fraud blacklist (`../data/fraud_blacklist.csv`)

Columns: `response_id, ip, child_name, delivery_email, reasons, source, added_at`.

It is written automatically every time `consent_management/send_survey_emails.py`
runs (including `--dry-run`), which screens each consent export and appends any
newly-flagged entries. The file is plain CSV and **hand-editable**: add a row
with `source = manual` to blacklist someone the automated screen missed, or
delete a row to clear a false positive. Re-runs upsert by `response_id` and
never drop or duplicate existing rows, so your manual edits survive. You can also
regenerate it directly with `python flag_suspicious_entries.py` in
`consent_management/`.

If the file does not exist yet (no consent screening has run), payment proceeds
with an empty blacklist and prints that no fraud list was found.

## Your manual workflow

1. Run `run_payments.bat` (or `run_payments_dryrun.bat` to preview).
2. Open `data\payment_report_unpaid.xlsx`. Buy + send $10 Amazon cards to the
   **Pay** sheet. Review the **Hold** sheet before deciding. **Do not pay the
   `Fraud (do not pay)` sheet** — those are blacklisted as bots/scammers.
3. As you send each card, mark `paid = yes` (and optionally `paid_date`) in
   `data\payment_tracker.xlsx`, then save.
4. Next run drops paid people from the report automatically.

## Policy switches (top of `manage_payments.py`)

- `HOLD_BUCKET_ENABLED` — `True` splits flagged completers into the Hold sheet.
  **Note:** `QUALITY_CONTROL.md` specifies the default protocol is to pay every
  completer and treat exclusion as analysis-only. Conditioning payment on
  quality flags deviates from that and should only be done with IRB sign-off.
  Set to `False` to put everyone payable in the Pay sheet (flags become
  review-only annotations).
- `HOLD_ON_ANY_FLAG` — **`True` (current setting, June 2026)** holds on any
  single flag (`n_flags >= 1`), the moderate-stringency option. `False` holds
  only on `exclude_recommended` (the filter's >= 2-flag combination rule). This
  is a payment-side split only; the shared `exclude_min_flags` (the analysis drop
  rule in `quality_config.json`) stays at 2, so the analysis repo is unaffected.

## Tuning stringency

`simulate_stringency.py` re-scores the latest export under several candidate
threshold sets (quality + fraud) and prints PASS/REJECT counts for both the raw
export and the completer population, changing nothing on disk:

```
python simulate_stringency.py            # newest export in ../ingest
python simulate_stringency.py --export PATH
```

The June 2026 stringency analysis, the Qualtrics comparison, and the rationale
for the applied moderate settings (reCAPTCHA 0.7, hold-on-any-flag) are written
up in **`STRINGENCY_ANALYSIS.md`** in this folder.

## quality_filter.py is vendored — do not edit it here

`quality_filter.py` and `quality_config.json` in this folder are **copies** of
the canonical files in the project root, pushed here by `sync_qc.bat`. Edit the
root copies, then re-run `sync_qc.bat` to update all three consumers. See
`QUALITY_CONTROL.md`.
