# examine_indv — look up one participant

A small **read-only** web tool for examining a single survey participant. Use it
when you want to understand one person (their answers, why they're being paid or
held), not when you want to run a batch. It writes to **nothing** — it only reads
the files the rest of the pipeline produces.

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

## How the data joins (important for future edits)

The join key everywhere is the **consent ResponseId**:

```
participant_tracker_auto.response_id  ==  payment_tracker.cid  ==  survey export "cid" column
```

Note the survey export's own `ResponseId` column is a *different* id and is **not**
the join key — use the embedded `cid` column (near the end of the export, around
column 144). All logic lives in `examine_data.py`; the server (`examine_app.py`)
is a thin stdlib `http.server` wrapper that serves JSON and one HTML page.

### Files read (all under the repo root, one level up)

| File | Used for |
|------|----------|
| `data/participant_tracker_auto.xlsx` (Participants sheet) | the roster / who exists |
| `data/payment_tracker.xlsx` (Payments sheet, by `cid`) | PAY/HOLD/FRAUD/PAID decision + reasons |
| `data/fraud_blacklist.csv` | fraud cross-check by RID, email, or IP |
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
