#!/usr/bin/env python3
"""
Free follow-up email agent for GitHub Actions – driven by emails.xlsx.

Every hour, for each row on the "Emails" sheet:
  PENDING  -> at Send Hour in the recipient's timezone, send Initial Message.
  AWAITING -> look for a reply from them (Gmail IMAP).
              reply       -> Status = REPLIED, stop.
              no reply    -> at Send Hour, if Interval Days elapsed and
                             Follow-ups Sent < Max Follow-ups, send Follow-up
                             Message in the same thread; else if cap hit -> DONE.
  REPLIED / DONE -> nothing.
The bot writes Status / First Sent / Last Sent / Follow-ups Sent / Reply Detected /
Last Checked / Next Action back into the sheet; the workflow commits it.

Secrets: GMAIL_ADDRESS, GMAIL_APP_PASSWORD.  DRY_RUN=1 logs without sending.
"""
import email, email.utils, imaplib, os, re, smtplib, sys
from datetime import date, datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

SHEET = Path(__file__).parent / "emails.xlsx"
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
DRY_RUN = os.environ.get("DRY_RUN") == "1"
DEFAULTS = {"Send Hour": 9, "Interval Days": 5, "Max Follow-ups": 3, "Timezone": "UTC"}
STAMP = "%Y-%m-%d %H:%M"


def log(m): print(f"[{datetime.now(timezone.utc):%Y-%m-%d %H:%M}Z] {m}", flush=True)


# ---------------------------------------------------------------- email I/O
def send_mail(to, subject, body, in_reply_to=None, references=None):
    msg = EmailMessage()
    msg["From"], msg["To"], msg["Subject"] = GMAIL_ADDRESS, to, subject
    msg["Date"] = email.utils.formatdate(localtime=True)
    mid = email.utils.make_msgid(domain="gmail.com")
    msg["Message-ID"] = mid
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = " ".join(references or [in_reply_to])
    msg.set_content(body)
    if DRY_RUN:
        log(f"DRY RUN – would send to {to}: {subject!r}"); return mid
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
        s.ehlo(); s.starttls(); s.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD); s.send_message(msg)
    log(f"SENT to {to}: {subject!r}"); return mid


def has_reply(from_addr, since: datetime, thread_ids):
    if DRY_RUN and not GMAIL_APP_PASSWORD: return False
    with imaplib.IMAP4_SSL("imap.gmail.com") as m:
        m.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        m.select('"[Gmail]/All Mail"', readonly=True)
        typ, data = m.search(None, "FROM", f'"{from_addr}"', "SINCE", since.strftime("%d-%b-%Y"))
        if typ != "OK" or not data or not data[0]: return False
        for num in data[0].split():
            typ, raw = m.fetch(num, "(BODY.PEEK[HEADER])")
            if typ != "OK": continue
            hdr = email.message_from_bytes(raw[0][1])
            d = email.utils.parsedate_to_datetime(hdr["Date"]) if hdr.get("Date") else None
            if d and d.tzinfo and d < since: continue
            return True  # any newer mail from them counts as a reply
    return False


# ---------------------------------------------------------------- sheet helpers
class Row:
    def __init__(self, ws, r, col):
        self.ws, self.r, self.col = ws, r, col
    def get(self, h, default=None):
        v = self.ws.cell(row=self.r, column=self.col[h]).value
        return default if v in (None, "") else v
    def set(self, h, v):
        self.ws.cell(row=self.r, column=self.col[h], value=v)


GRACE_HOURS = 2   # a late/skipped cron run may still send up to 2h after the slot


def parse_hours(v) -> list[int]:
    """'9, 15' / '9;15' / '9 15' / 9 / 9.15 (Excel comma-decimal) -> [9, 15]"""
    if isinstance(v, float) and v.is_integer(): v = int(v)
    hours = sorted({int(x) for x in re.findall(r"\d+", str(v))})
    if not hours or any(h > 23 for h in hours):
        raise ValueError(f"Send Hour must be 0-23, e.g. 9 or '9, 15' (got {v!r})")
    return hours


def day_allowed(day: date, last_sent: datetime | None, interval: int) -> bool:
    if last_sent is None: return True
    d = (day - last_sent.date()).days
    return d == 0 or d >= interval          # same send-day, or interval elapsed


def due_slot(now: datetime, hours, last_sent, interval):
    """Return the slot datetime if a send is due right now, else None."""
    passed = [h for h in hours if h <= now.hour]
    if not passed: return None
    slot = now.replace(hour=passed[-1], minute=0, second=0, microsecond=0)
    if (now - slot).total_seconds() > GRACE_HOURS * 3600: return None
    if last_sent and last_sent >= slot: return None
    return slot if day_allowed(now.date(), last_sent, interval) else None


def next_slot(now: datetime, hours, last_sent, interval):
    for add in range(0, interval + 2):
        day = now.date() + timedelta(days=add)
        if not day_allowed(day, last_sent, interval): continue
        for h in hours:
            cand = now.replace(year=day.year, month=day.month, day=day.day,
                               hour=h, minute=0, second=0, microsecond=0)
            if cand > now and (last_sent is None or cand > last_sent):
                return cand
    return None


def fmt_next(label, now, hours, last_sent, interval, tzname):
    n = next_slot(now, hours, last_sent, interval)
    if not n: return label
    when = "today" if n.date() == now.date() else ("tomorrow" if (n.date() - now.date()).days == 1 else n.strftime("%a %d %b"))
    return f"{label} {when} at {n:%H:%M} {tzname}"


def process(row: Row):
    to = str(row.get("Email", "")).strip()
    if not to or "@" not in to: return
    status = str(row.get("Status", "PENDING")).strip().upper()
    if status in ("REPLIED", "DONE"): return

    tz_raw = str(row.get("Timezone", DEFAULTS["Timezone"])).strip()
    m = re.search(r"\(([^)]+)\)\s*$", tz_raw)   # "Sweden (Europe/Stockholm)" -> Europe/Stockholm
    tzname = m.group(1) if m else tz_raw
    try: tz = ZoneInfo(tzname)
    except Exception:
        row.set("Next Action", f"ERROR: unknown timezone '{tz_raw}'"); return
    now_local = datetime.now(tz)
    row.set("Last Checked", datetime.now(timezone.utc).strftime(STAMP) + " UTC")
    try: hours = parse_hours(row.get("Send Hour", DEFAULTS["Send Hour"]))
    except ValueError as e:
        row.set("Next Action", f"ERROR: {e}"); return
    interval = max(int(row.get("Interval Days", DEFAULTS["Interval Days"])), 1)
    max_fu = int(row.get("Max Follow-ups", DEFAULTS["Max Follow-ups"]))
    name = str(row.get("Name", "") or "")
    subject = str(row.get("Subject", "") or "").strip()
    ids = [x for x in str(row.get("Message IDs", "") or "").split() if x]
    fu_sent = int(row.get("Follow-ups Sent", 0) or 0)
    ls = row.get("Last Sent")
    last_sent = datetime.strptime(str(ls), STAMP).replace(tzinfo=tz) if ls else None

    # --- initial ---
    if status == "PENDING":
        if not subject or not row.get("Initial Message"):
            row.set("Next Action", "ERROR: Subject and Initial Message required"); return
        if not due_slot(now_local, hours, None, interval):
            row.set("Next Action", fmt_next("Send initial email", now_local, hours, None, interval, tzname))
            return
        mid = send_mail(to, subject, str(row.get("Initial Message")).format(name=name))
        stamp = now_local.strftime(STAMP)
        row.set("Status", "AWAITING"); row.set("First Sent", stamp); row.set("Last Sent", stamp)
        row.set("Follow-ups Sent", 0); row.set("Message IDs", mid)
        row.set("Next Action", fmt_next("Follow-up #1", now_local, hours, now_local, interval, tzname))
        return

    # --- awaiting: reply check ---
    first_sent = datetime.strptime(str(row.get("First Sent")), STAMP).replace(tzinfo=tz)
    if has_reply(to, first_sent, ids):
        row.set("Status", "REPLIED"); row.set("Reply Detected", now_local.strftime(STAMP))
        row.set("Next Action", "Reply received – no more emails"); log(f"{to}: REPLY"); return

    if fu_sent >= max_fu:
        row.set("Status", "DONE"); row.set("Next Action", "Max follow-ups reached, no reply"); return

    if not due_slot(now_local, hours, last_sent, interval):
        row.set("Next Action", fmt_next(f"Follow-up #{fu_sent+1}", now_local, hours, last_sent, interval, tzname))
        return

    subj = subject if subject.lower().startswith("re:") else f"Re: {subject}"
    mid = send_mail(to, subj, str(row.get("Follow-up Message") or "").format(name=name),
                    in_reply_to=ids[-1] if ids else None, references=ids)
    ids.append(mid); fu_sent += 1
    row.set("Message IDs", " ".join(ids)); row.set("Follow-ups Sent", fu_sent)
    row.set("Last Sent", now_local.strftime(STAMP))
    row.set("Next Action", "Max follow-ups reached after this" if fu_sent >= max_fu
            else fmt_next(f"Follow-up #{fu_sent+1}", now_local, hours, now_local, interval, tzname))


def main():
    if not DRY_RUN and not (GMAIL_ADDRESS and GMAIL_APP_PASSWORD):
        log("ERROR: GMAIL_ADDRESS / GMAIL_APP_PASSWORD not set"); return 1
    wb = load_workbook(SHEET)          # keeps formulas on Dashboard
    ws = wb["Emails"]
    col = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    for r in range(2, ws.max_row + 1):
        try: process(Row(ws, r, col))
        except Exception as e:
            log(f"row {r}: ERROR {e!r}")
            ws.cell(row=r, column=col["Next Action"], value=f"ERROR: {e}")
    wb.save(SHEET); return 0


if __name__ == "__main__":
    sys.exit(main())
