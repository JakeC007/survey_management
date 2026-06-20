# Throwaway-domain review — consent export (2026-06-18)

Read-only analysis. Nothing in the screener changed.

## Files in this package

- `email_throwaway_domains.txt` — 4239 domains (public maintained list + 4 curated extras).
- `email_domain_whitelist.txt` — schools/.edu to never flag.
- `domain_signals.csv` — every non-mainstream domain in the export with its signals.

## Domains seen in YOUR export, with recommendation

`new` = rows not already caught by the current screener.

| domain | rows | new | distinct IPs | reCAPTCHA mean | on public list | recommendation |
|---|---|---|---|---|---|---|
| haren.uk | 6 | 6 | 5 | 0.67 | Y | disposable (on public list) |
| choco.la | 4 | 1 | 4 | 0.77 | Y | disposable (on public list) |
| instaddr.win | 4 | 1 | 3 | 0.68 | Y | disposable (on public list) |
| meruado.uk | 4 | 3 | 4 | 0.70 | Y | disposable (on public list) |
| cps.edu | 3 | 3 | 3 | 1.00 | - | WHITELIST (real school) |
| digdig.org | 3 | 3 | 3 | 0.93 | Y | disposable (on public list) |
| ichigo.me | 3 | 0 | 2 | 0.70 | Y | disposable (on public list) |
| instaddr.ch | 3 | 2 | 3 | 0.80 | Y | disposable (on public list) |
| owleyes.ch | 3 | 3 | 3 | 0.87 | - | ADD to list (curated) |
| simaenaga.com | 3 | 2 | 2 | 0.77 | - | ADD to list (curated) |
| ucls.uchicago.edu | 3 | 3 | 3 | 0.77 | - | WHITELIST (real school) |
| exdonuts.com | 2 | 1 | 2 | 0.70 | Y | disposable (on public list) |
| gmil.com | 2 | 2 | 2 | 0.95 | - | HOLD: typo-squat of gmail.com - treat as malformed/typo, not disposable (could be a real user's slip) |
| hamham.uk | 2 | 2 | 1 | 0.85 | Y | disposable (on public list) |
| instaddr.uk | 2 | 1 | 2 | 0.75 | Y | disposable (on public list) |
| kpay.be | 2 | 1 | 2 | 0.75 | Y | disposable (on public list) |
| kpost.be | 2 | 0 | 2 | 0.70 | Y | disposable (on public list) |
| lsoc.org | 2 | 2 | 2 | 1.00 | - | HOLD: could be a legitimate org; reCAPTCHA 1.0, weak signal - confirm before listing |
| send4.uk | 2 | 0 | 2 | 0.75 | - | ADD to list (curated) |
| via.tokyo.jp | 2 | 0 | 2 | 0.35 | - | ADD to list (curated) |
| 599rocks.com | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| addrin.uk | 1 | 1 | 1 | 0.70 | Y | disposable (on public list) |
| bangban.uk | 1 | 0 | 1 | 0.60 | Y | disposable (on public list) |
| boxfi.uk | 1 | 1 | 1 | 0.90 | Y | disposable (on public list) |
| cox.net | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| delatorre.tv | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| eay.jp | 1 | 0 | 1 | 0.60 | Y | disposable (on public list) |
| f5.si | 1 | 1 | 1 | 0.80 | - | leave (mainstream-ish / low signal) |
| fuwa.li | 1 | 1 | 1 | 1.00 | Y | disposable (on public list) |
| fuwamofu.com | 1 | 1 | 1 | 0.50 | Y | disposable (on public list) |
| goldercollegeprep.org | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| hotmail.cm | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| instmail.uk | 1 | 1 | 1 | 0.90 | Y | disposable (on public list) |
| mbox.re | 1 | 0 | 1 | 0.90 | Y | disposable (on public list) |
| merry.pink | 1 | 1 | 1 | 0.90 | Y | disposable (on public list) |
| moimoi.re | 1 | 0 | 1 | 0.40 | - | leave (mainstream-ish / low signal) |
| otona.uk | 1 | 1 | 1 | 0.80 | - | leave (mainstream-ish / low signal) |
| quicksend.ch | 1 | 0 | 1 | 0.90 | - | leave (mainstream-ish / low signal) |
| rpschools.net | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| sendapp.uk | 1 | 1 | 1 | 0.90 | - | leave (mainstream-ish / low signal) |
| stayhome.li | 1 | 1 | 1 | 0.80 | - | leave (mainstream-ish / low signal) |
| svk.jp | 1 | 1 | 1 | 0.90 | - | leave (mainstream-ish / low signal) |
| uchicagocharter.org | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |
| usako.net | 1 | 0 | 1 | 0.40 | - | leave (mainstream-ish / low signal) |
| wi.rr.com | 1 | 1 | 1 | 1.00 | - | leave (mainstream-ish / low signal) |

## Notes

- Most of these throwaway domains have reCAPTCHA means of 0.7–1.0, so even the tightened `recaptcha_min: 0.7` rule (raised from 0.5 in June 2026) still misses the ones at 0.8–1.0. A domain check catches farms that pass reCAPTCHA.
- Held back pending your call:
  - `gmil.com` — typo-squat of gmail.com - treat as malformed/typo, not disposable (could be a real user's slip)
  - `lsoc.org` — could be a legitimate org; reCAPTCHA 1.0, weak signal - confirm before listing
- Suggested wiring (when you decide): add a TECHNICAL soft flag `flag_throwaway_email` that trips when the delivery domain is in the list AND not in the whitelist. Soft, so it only excludes alongside a second signal — keeps a one-off real user on an odd domain safe.
