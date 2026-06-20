# dashboard

A local web console for the survey pipeline. It shows live numbers from the
shared tracker and gives you one button per step that runs the **existing**
scripts in `consent_management/`, `response_management/`, and
`payment_management/`. It does not duplicate any pipeline logic — every button
just launches the same `.bat` / `.py` you already run by hand.

## Run it

1. Make sure you've run `setup.bat` in `survey_management/` once (creates the
   shared `.venv`).
2. Copy the config template and fill in your survey links:
   ```
   copy config.sample.yaml config.yaml
   ```
3. Double-click **`run_dashboard.bat`**.
4. Your browser opens to `http://127.0.0.1:5000`. Leave the console window open;
   close it (or Ctrl+C) to stop.

If `config.yaml` is missing, the dashboard falls back to `config.sample.yaml`
(with placeholder links) and prints a reminder to copy it.

## Files

- `app.py` — the local server: reads the tracker, serves the page, runs the scripts.
- `index.html` — the page template (`app.py` fills in the live numbers).
- `config.sample.yaml` — committed template. Copy to `config.yaml`.
- `config.yaml` — your real settings and survey links (gitignored).
- `run_dashboard.bat` — double-click launcher.

Windows + classic Outlook only, same as the rest of the pipeline. The send and
draft buttons drive Outlook exactly like the `.bat` files do.

## What the buttons do

| Step | Button | Runs |
|---|---|---|
| 1 | Dry run | `send_survey_emails.py --dry-run` |
| 1 | Draft invites | `consent_management\run_draft.bat` |
| 1 | Send invites | `consent_management\run.bat` |
| 2 | Dry run | `response_management\run_responses_dryrun.bat` |
| 2 | Draft reminders | `response_management\run_responses_draft.bat` |
| 2 | Send reminders | `response_management\run_responses.bat` |
| 3 | Dry run | `payment_management\run_payments_dryrun.bat` |
| 3 | Build report | `payment_management\run_payments.bat` |
| 3 | Open unpaid report | opens `data\payment_report_unpaid.xlsx` |

Send buttons ask for confirmation first. Output (the script's console text)
appears in the Output panel at the bottom of the page.

Drop the correct Qualtrics export ZIP in `..\ingest\` before running a step:
the consent ZIP for step 1, the teen response ZIP for steps 2 and 3.

## Stats shown

Read live from `..\data\participant_tracker_auto.xlsx` (and the payment files):

- **IRB consent received** — rows where BOTH the parent consent and the teen
  assent cleared (affirmations + signatures). Inferred from `status`/`reason`:
  excludes `INCOMPLETE` and `SUSPICIOUS`, and any row whose reason cites a
  consent, assent, or signature failure. (A row that is otherwise valid but was
  marked ineligible only for a bad delivery email still counts here, since both
  parties did consent.)
- **Survey invites sent** — rows with `emailed_at` set
- **Surveys completed** — rows with `survey_completed_at` set
- **Gift cards paid** — `paid` rows in `payment_tracker.xlsx`
- **Ineligible for payment** — rows with status `INELIGIBLE` or `INCOMPLETE`
- **Flagged as fraud** — rows with status `SUSPICIOUS` (bot/fraud/duplicate
  entries the consent screen refused to email), tracked separately from
  ineligibility
- **Pay / Hold / Fraud** — sheet counts from `payment_report_unpaid.xlsx`. The
  Fraud tag on the payment card is the number of completers whose `cid` or survey
  IP is on the blacklist and who will not be paid (distinct from Hold). The Hold
  count also includes completers whose delivery email is on a known
  throwaway/disposable domain, held for review rather than auto-paid.

## File locking / running things at once

The dashboard guards against clashing with the scripts over the shared xlsx
files:

- Only one step runs at a time. Starting a step while another is running is
  rejected with a message — no two scripts write the tracker at once.
- While a step is running, the stats panel stops reading the xlsx files and
  shows the last snapshot, so a page refresh can't open a file mid-write. Fresh
  numbers appear on the next refresh after the run finishes.
- The reader opens files read-only and closes them immediately, and any read
  error is swallowed (numbers just show stale until the next refresh).

The one thing it can't control: **don't have the tracker or payment files open
in Excel while a step runs.** Excel takes an exclusive lock, and the script will
fail to write. Close the workbook before running, reopen it after.

## config.yaml

`config.yaml` is gitignored (the repo's `.gitignore` ignores `config.yaml`
everywhere), so your real survey links stay off GitHub. Only
`config.sample.yaml` is committed. Copy the sample to `config.yaml` and edit
that. It holds:

- `qualtrics_data_sources` — links to the two Qualtrics surveys you export from,
  shown as one-click shortcuts at the top of the page. **Paste the consent
  survey's Data & Analysis URL** where marked (the teen survey is prefilled from
  `..\config.yaml`).
- `paths` — where to read the tracker and payment files (defaults match the
  shared layout).
- `server` — host, port, and whether to auto-open the browser.
