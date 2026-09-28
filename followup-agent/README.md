# Follow-up email agent (free, GitHub Actions + Excel tracker)

`emails.xlsx` is the whole control panel. You fill the **yellow** columns;
the bot fills the **grey** ones every hour and commits the file back, so the
sheet on GitHub always shows what was sent, when, and what happens next.

## Files
| File | What |
|---|---|
| `emails.xlsx` | Emails sheet (your rows + bot status), Dashboard (counts), Timezones (dropdown list) |
| `followup.py` | The agent |
| `.github/workflows/followup.yml` | Hourly schedule + commit-back |
| `make_sheet.py` | Regenerates a blank `emails.xlsx` if you ever break it |

## Setup (10 min)
1. Create a **private** GitHub repo (the sheet contains addresses) and push these files. Private repos get 2,000 free Actions minutes/month; this uses ~720.
2. Gmail: Google Account → Security → 2-Step Verification on → search "App passwords" → create one.
3. Repo → Settings → Secrets and variables → Actions → add `GMAIL_ADDRESS` and `GMAIL_APP_PASSWORD`.
4. Open `emails.xlsx`, fill the yellow columns (row 2 is an example – overwrite it), commit, push.
5. Test: uncomment `DRY_RUN: "1"` in the workflow, Actions → *Run workflow*, read the log and the updated sheet. Re-comment, push.

## Columns you fill (yellow)
Email · Name · Timezone (dropdown: Switzerland, Denmark, Sweden, Netherlands, Singapore, all 6 USA zones, plus others – or type any IANA name) · Send Hour (one or more hours 0–23 in *their* local time: `9` = once, `9, 15` = twice, `9, 13, 17` = three times per send day) · Subject · Initial Message · Follow-up Message · Interval Days · Max Follow-ups.
`{name}` inside a message is replaced by the Name column. Blank Send Hour / Interval / Max default to 9 / 5 / 3.

## Columns the bot fills (grey)
Status (`PENDING → AWAITING → REPLIED` or `DONE`) · First Sent · Last Sent · Follow-ups Sent · Reply Detected · Last Checked · Next Action (plain-English plan, or an `ERROR:` message if a row is broken).

- Add a person: new row, push.
- Restart a person: clear their grey cells, Status → `PENDING`, push.
- Stop a person: Status → `DONE`.
- **Don't edit the sheet on GitHub while the bot may be committing** – pull first, edit, push. The workflow has a concurrency lock so two runs never collide, but a stale local copy will conflict on push.

## Sending several times a day
`Send Hour = 9, 15` with `Interval Days = 1` sends at 09:00 and 15:00 their time, every day, until they reply.
`Interval Days = 2` with the same hours sends twice on every second day. Each send counts as one follow-up, so raise **Max Follow-ups** to match (twice daily for a week ≈ 14). If a cron run is late, the bot still sends within 2 hours of the slot; later than that, the slot is skipped rather than sent at an odd hour.

## Limitations
- Reply detection = any email from that address after your first send. A reply from a different address (assistant, personal) is missed – glance at your inbox before trusting `AWAITING`.
- GitHub cron can slip 5–30 min. The "already sent today" guard means it still sends once, just late.
- Scheduled workflows are disabled after 60 days of no commits; the bot's own commits keep it alive while anything is active.
- Dashboard formulas show blank in GitHub's file preview (openpyxl doesn't cache values) – they compute when you open the file in Excel/LibreOffice.

## Cadence – read this
Daily follow-ups to professors get you filtered or ignored. Norms that get replies: first nudge after 7–10 days, second after another 7–14, stop at 2–3 total. Defaults are 5 days / max 3; daily is possible (`Interval Days = 1`) but not recommended for outreach.
