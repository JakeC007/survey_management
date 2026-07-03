# examine_indv — look up one participant

A small web tool for examining a single survey participant. Use it when you want
to understand one person (their answers, why they're being paid or held), not
when you want to run a batch. It also has a **Review queue** for stepping through
every flagged participant one at a time (see below).

It is read-only **except** for two explicit actions: the "Mark as fraud" button
and the "Clear / keep" button in the review queue. Both are described below.

This tool is for one person at a time. To actually **pay** completed
participants in a batch — copy their emails into Amazon, mark them paid, and send
the thank-you email — use the payment console `payment_management/pay_app.py`
(`run_pay_app.bat`). If you mark someone paid there and later mark them fraud
here, the fraud flag stands; review before clawing back an already-sent card.

## Run it

```
python examine_indv/examine_app.py
```

(or `examine_indv\examine_app.py` on Windows, using the project `.venv`). It
loads the data, opens your browser to `http://127.0.0.1:8765`, and prints the
URL. Pass a port number as an argument to change it: `python examine_app.py 9000`.

Then search by:

- **Name** — first only (`Eric`) or first + last (`Harper Patrick`). Matching is
  by word-prefix on both the parent and child name, so partials work (`Har`) and
  every match is listed when a name collides.
- **Response / RID** — `R_3Q7zMYZjRHFPI2d` (anything matching `^R_…`).
- **Email** — substring match on `delivery_email`.

If exactly one person matches, it opens straight to their detail; multiple
matches show a pick-list; no match shows a friendly "no one found".

## What it shows

1. **Status** — one of `PAY`, `HOLD`, `FRAUD`, `PAID`, or `NOT EVALUATED`, with
   the reason underneath.
2. **Completion** — if there's no completed survey, it says so and shows no
   answers (status still appears).
3. **Key metadata** — quality signals (flag count/reasons, exclude flag,
   throwaway email, consent, IP) plus survey signals (duration, reCAPTCHA score,
   straightlining %, unanswered %, finished, duplicate-respondent).
4. **Answers** — the substantive Q&A with full question text. Timing columns,
   lat/long, and system fields are filtered out.

## Mark as fraud (the one write action)

The detail view has a **Mark as fraud** button. It opens a confirmation popup
(with an optional reason), and on confirm it writes the person into every place
the project tracks fraud, so the change is durable and the dashboard reflects it:

1. **`data/fraud_blacklist.csv`** — appends a row (`source=manual`). This is the
   source of truth: the payment pipeline rebuilds its fraud verdict from this
   file on every run, so the mark survives re-runs.
2. **`data/payment_tracker.xlsx`** — sets `fraud=yes`, `fraud_reason`, and
   `exclude_recommended=yes` on the person's row.
3. **`data/payment_report_unpaid.xlsx`** — removes them from the Pay/Hold sheets
   and adds them to the Fraud sheet, so the dashboard's Fraud count updates
   immediately without waiting for a pipeline run.

It is **atomic** (temp file then replace) and **idempotent** (re-marking someone
already fraud is a no-op per file, and won't stack reasons). If a workbook is
open in Excel the write fails with a clear message instead of corrupting it, so
close `payment_tracker.xlsx` / `payment_report_unpaid.xlsx` before marking. All
write logic lives in `examine_write.py`; nothing else in this folder writes.

Note it does **not** change `participant_tracker_auto.xlsx` status. That column
is owned by the consent/response pipeline, and the dashboard's separate
"Flagged as fraud" card counts consent-time bot detections, which is a different
thing from a manual payment-side fraud mark.

## Review flagged participants (the Back / Next queue)

The link at the top of the search page (or `/review` directly) opens a review
queue for working through everyone who carries quality flags, one at a time,
instead of searching for them by name. Use it when you want to clear a backlog of
flagged responses and make a fraud-or-keep call on each.

**Who's in the queue.** Every participant with `n_flags > 0` in
`payment_tracker.xlsx`, **minus** anyone already marked `FRAUD` and anyone you've
already cleared here. Most-flagged appear first; a "hide already-paid" toggle
drops people who have already been paid. The queue is rebuilt from disk on every
load, so it always reflects the current sheets.

**Navigation.** Back / Next buttons (or the ← / → arrow keys) move between items.
A counter and progress bar show where you are. Each item shows the same status,
reason, metadata, and answers as the search detail view, plus the specific
`flag_reasons` that put the person in the queue.

**Two decisions per item:**

- **Mark as fraud** — identical to the search-view button above; routes through
  `examine_write.mark_fraud` and writes the three fraud files. The person leaves
  the queue (they're now `FRAUD`).
- **Clear / keep** — judges the person legitimate and moves them from **Hold** to
  **Pay**, so the dashboard's Hold count drops and Pay count rises. Like
  Mark-as-fraud, it writes to every place that matters:

  1. **`data/payment_tracker.xlsx`** — sets `exclude_recommended=no` and stamps a
     `kept via examine_indv review` note on the row, so they read as `PAY`.
  2. **`data/payment_report_unpaid.xlsx`** — moves the row from the
     `Hold (flagged)` sheet to the `Pay (no hold)` sheet, so the dashboard counts
     (which are sheet sizes) update immediately, without a pipeline run.
  3. **`data/review_state.csv`** — records the keep (`decision=cleared`), which
     drops the person from this queue **and** makes the next `manage_payments.py`
     run keep them in Pay (see below).

  Marking does not touch their `fraud` flag and never moves anyone *into* Hold.

**Durable across pipeline re-runs.** `manage_payments.py` recomputes the Hold/Pay
split from quality flags every run, so a one-off edit to `payment_tracker.xlsx`
would be undone on the next build. To make a keep stick, `manage_payments.py`
reads `review_state.csv` at startup (`load_kept_cids`) and forces those cids out
of Hold in `is_held()` — the same way the fraud blacklist forces cids into Fraud.
So a kept person stays in Pay no matter how many times the pipeline re-runs.

`review_state.csv` is **only read by the review queue and `manage_payments.py`,
and only written here**, so a clear's CSV write can never collide with an open
Excel workbook. The two xlsx writes reuse `examine_write`'s atomic save and the
same Excel-lock guard as Mark-as-fraud: if a workbook is open in Excel the keep
is **not** half-applied — nothing is recorded and you get a clear "close it and
retry" message. Re-clearing an already-cleared person is a no-op. All queue and
clear logic lives in `examine_review.py`; the server reuses the existing `STORE`
and the same `/api/detail` and `/api/mark_fraud` endpoints.

## How the data joins (important for future edits)

The join key everywhere is the **consent ResponseId**:

```
participant_tracker_auto.response_id  ==  payment_tracker.cid  ==  survey export "cid" column
```

Note the survey export's own `ResponseId` column is a *different* id and is **not**
the join key — use the embedded `cid` column (near the end of the export, around
column 144). Read logic lives in `examine_data.py`, the fraud writer in
`examine_write.py`, and the review queue + clear logic in `examine_review.py`.
The server (`examine_app.py`) is a thin stdlib `http.server` wrapper that serves
JSON and two HTML pages (the search page and `/review`).

### Files read (all under the repo root, one level up)

| File | Used for |
|------|----------|
| `data/participant_tracker_auto.xlsx` (Participants sheet) | the roster / who exists |
| `data/payment_tracker.xlsx` (Payments sheet, by `cid`) | PAY/HOLD/FRAUD/PAID decision + reasons |
| `data/fraud_blacklist.csv` | fraud cross-check by RID, email, or IP |
| `data/review_state.csv` | "cleared / keep" decisions from the review queue (created on first clear) |
| newest `ingest/K12 *.zip` (not the `Consent …` zips) | the actual answers |

The newest survey ZIP is chosen by parsing the date in the filename
(`…_June 25, 2026_06.18.zip`), falling back to file mtime. Consent ZIPs are
skipped by name.

### Status logic (precedence)

1. **FRAUD** — `payment_tracker.fraud == yes` **or** a hit in `fraud_blacklist.csv`
   (matched on RID / email / IP). Reason = `fraud_reason` and/or the blacklist
   `reasons`.
2. **PAID** — `paid == yes` (with date/amount if present).
3. **HOLD** — `exclude_recommended == yes`. Reason = `flag_reasons`.
4. **PAY** — a payment-tracker row exists with no exclusion. Minor flags noted.
5. **NOT EVALUATED** — no payment-tracker row yet; falls back to the tracker's
   own pipeline status (SUSPICIOUS → HOLD, INCOMPLETE/INELIGIBLE → not evaluated).

This mirrors the payment pipeline's buckets; if those rules change in
`payment_management/`, update `_status()` in `examine_data.py` to match.

## Dependencies

Standard library only, plus `openpyxl` (already in the project `requirements.txt`).
No web framework, no extra install.
