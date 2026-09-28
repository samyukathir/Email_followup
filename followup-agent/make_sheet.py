"""Creates emails.xlsx – the single sheet you edit and the bot updates."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

INPUT = [  # (header, width, note)
    ("Email", 30, "Recipient address (required)"),
    ("Name", 22, "Used for {name} in messages"),
    ("Timezone", 32, "Pick from dropdown or type an IANA zone"),
    ("Send Hour", 13, "One or more hours 0-23 in their local time, e.g. 9 or 9, 15"),
    ("Subject", 40, "First email subject; follow-ups get 'Re:'"),
    ("Initial Message", 60, "Body of the first email. {name} is replaced."),
    ("Follow-up Message", 60, "Body of each follow-up (sent in same thread)"),
    ("Interval Days", 13, "Days between sends. 1 = daily"),
    ("Max Follow-ups", 14, "Stop after this many follow-ups"),
]
BOT = [
    ("Status", 13, "PENDING → AWAITING → REPLIED / DONE"),
    ("First Sent", 19, "When the initial email went out"),
    ("Last Sent", 19, "Most recent send"),
    ("Follow-ups Sent", 15, "Count so far"),
    ("Reply Detected", 19, "When the bot saw their reply"),
    ("Last Checked", 19, "Last time the bot looked at this row"),
    ("Next Action", 45, "What the bot plans to do next"),
    ("Message IDs", 10, "internal – for threading (hidden)"),
]
HEADERS = [h for h, _, _ in INPUT + BOT]
FONT = "Arial"
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
ws = wb.active
ws.title = "Emails"

for i, (h, w, _) in enumerate(INPUT + BOT, start=1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = Font(name=FONT, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F4E78" if i <= len(INPUT) else "595959")
    c.alignment = Alignment(vertical="center", wrap_text=True)
    c.border = border
    ws.column_dimensions[get_column_letter(i)].width = (40 if h == "Timezone" else w)
ws.row_dimensions[1].height = 30
ws.freeze_panes = "B2"

# example row
example = ["professor@university.edu", "Prof. Example", "Switzerland (Europe/Zurich)", "9, 15",
           "Research placement inquiry – AI for biomedicine",
           "Dear {name},\n\nI'm Sam, a developer working on ... I'd love to discuss a research placement in your lab in 2027.\n\nBest regards,\nSam",
           "Dear {name},\n\nGently following up on my note below in case it got buried. Happy to send a CV or project summary.\n\nBest regards,\nSam",
           5, 3, "PENDING"]
for i, v in enumerate(example, start=1):
    c = ws.cell(row=2, column=i, value=v)
    c.font = Font(name=FONT, color="0000FF" if i <= len(INPUT) else "000000")
    c.alignment = Alignment(vertical="top", wrap_text=True)
    c.border = border
# input fill for editable rows
for r in range(2, 202):
    for i in range(1, len(INPUT) + 1):
        ws.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FFF9E6")
        ws.cell(row=r, column=i).font = Font(name=FONT, color="0000FF")
        ws.cell(row=r, column=i).alignment = Alignment(vertical="top", wrap_text=True)
    for i in range(len(INPUT) + 1, len(HEADERS) + 1):
        ws.cell(row=r, column=i).fill = PatternFill("solid", fgColor="EDEDED")
        ws.cell(row=r, column=i).font = Font(name=FONT)
        ws.cell(row=r, column=i).alignment = Alignment(vertical="top", wrap_text=True)
ws.column_dimensions[get_column_letter(len(HEADERS))].hidden = True

# Timezones sheet + dropdown
tz = wb.create_sheet("Timezones")
zones = [
 # --- priority countries ---
 "Switzerland (Europe/Zurich)",
 "Denmark (Europe/Copenhagen)",
 "Sweden (Europe/Stockholm)",
 "Netherlands (Europe/Amsterdam)",
 "Singapore (Asia/Singapore)",
 "USA – Eastern (America/New_York)",
 "USA – Central (America/Chicago)",
 "USA – Mountain (America/Denver)",
 "USA – Arizona (America/Phoenix)",
 "USA – Pacific (America/Los_Angeles)",
 "USA – Alaska (America/Anchorage)",
 "USA – Hawaii (Pacific/Honolulu)",
 # --- others ---
 "UK (Europe/London)", "Ireland (Europe/Dublin)", "Germany (Europe/Berlin)", "France (Europe/Paris)",
 "Belgium (Europe/Brussels)", "Austria (Europe/Vienna)", "Norway (Europe/Oslo)", "Finland (Europe/Helsinki)",
 "Spain (Europe/Madrid)", "Italy (Europe/Rome)", "India (Asia/Kolkata)", "Japan (Asia/Tokyo)",
 "UTC (UTC)"]
tz["A1"] = "Country (IANA time zone)"; tz["A1"].font = Font(name=FONT, bold=True)
for i, z in enumerate(zones, start=2):
    tz.cell(row=i, column=1, value=z).font = Font(name=FONT)
tz.column_dimensions["A"].width = 40
dv = DataValidation(type="list", formula1=f"=Timezones!$A$2:$A${len(zones)+1}",
                    allow_blank=True, errorStyle="warning",
                    errorTitle="Not in list", error="OK if it is a valid IANA zone, e.g. Asia/Bangkok.")
ws.add_data_validation(dv); dv.add("C2:C201")
for r in range(2, 202):
    ws.cell(row=r, column=4).number_format = "@"   # text, so "9, 15" isn't read as a decimal

# Dashboard
db = wb.create_sheet("Dashboard", 0)
db["A1"] = "Follow-up agent – status"; db["A1"].font = Font(name=FONT, bold=True, size=14)
rows = [("Total rows with an email", '=COUNTA(Emails!A2:A201)'),
        ("PENDING (not yet sent)", '=COUNTIF(Emails!J2:J201,"PENDING")'),
        ("AWAITING reply", '=COUNTIF(Emails!J2:J201,"AWAITING")'),
        ("REPLIED", '=COUNTIF(Emails!J2:J201,"REPLIED")'),
        ("DONE (cap reached, no reply)", '=COUNTIF(Emails!J2:J201,"DONE")'),
        ("Follow-ups sent in total", '=SUM(Emails!M2:M201)')]
for i, (label, f) in enumerate(rows, start=3):
    db.cell(row=i, column=1, value=label).font = Font(name=FONT)
    db.cell(row=i, column=2, value=f).font = Font(name=FONT, bold=True)
db.column_dimensions["A"].width = 34; db.column_dimensions["B"].width = 10
db["A11"] = "How to use"; db["A11"].font = Font(name=FONT, bold=True)
legend = [
 "1. Fill the yellow columns (A–I) on the Emails sheet – one row per person. Row 2 is an example; edit or delete it.",
 "2. Leave Status blank or PENDING. The bot fills the grey columns (J–Q) every hour and commits the file back.",
 "3. To add someone later: add a row, commit, push. To restart someone: clear their grey cells and set Status to PENDING.",
 "4. To stop someone manually: set Status to DONE.",
 "5. Do not rename headers or sheets – the bot reads them by name.",
 "Several sends per day: put multiple hours in Send Hour, e.g. 9, 15 (twice a day) or 9, 13, 17. Raise Max Follow-ups to match.",
 "Cadence note: daily nudges to professors get you ignored. 7+ days and max 2–3 follow-ups is the norm.",
]
for i, t in enumerate(legend, start=12):
    db.cell(row=i, column=1, value=t).font = Font(name=FONT)
wb.save("emails.xlsx")
print("emails.xlsx written")
