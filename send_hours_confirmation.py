"""
Weekly Hours Confirmation Emailer
----------------------------------
Pulls time entries from a Smartsheet helper sheet, groups them by employee,
builds a clean HTML table (Date / Job / Hours + weekly Total), and emails
each employee their week's hours for confirmation before payroll.

Setup:
    pip install smartsheet-python-sdk

Required environment variables (set as secrets if running in GitHub Actions):
    SMARTSHEET_API_TOKEN   - your Smartsheet API access token
    GMAIL_ADDRESS           - the Workspace address sending the emails
    GMAIL_APP_PASSWORD      - App Password generated for that Gmail account
                              (Google Account > Security > 2-Step Verification > App Passwords)
"""

import os
import smtplib
from collections import defaultdict
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import smartsheet

# ---------------------------------------------------------------------------
# CONFIG - edit these to match your helper sheet
# ---------------------------------------------------------------------------

SMARTSHEET_API_TOKEN = os.environ["SMARTSHEET_API_TOKEN"]
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

HELPER_SHEET_ID = 6881750327185284

# TODO: these must match your helper sheet's column titles exactly
COL_EMPLOYEE_NAME = "Name"
COL_EMPLOYEE_EMAIL = "Email"
COL_DATE = "Date"
COL_JOB = "Customer Job"
COL_HOURS = "Hours"
COL_NOTES = "Notes"

# Pay week runs Monday - Sunday. By default the script auto-calculates the
# most recently COMPLETED pay week based on today's date, so triggering it on
# a Monday or Tuesday still correctly pulls last week (not the in-progress one).
#
# To manually pull a specific week instead (e.g. re-sending after a correction),
# set this to that week's Monday date as "YYYY-MM-DD". Leave as None for auto.
# Can also be set via the MANUAL_WEEK_START env var (used by the GitHub Actions
# workflow_dispatch input) - env var takes precedence if present.
MANUAL_WEEK_START = os.environ.get("MANUAL_WEEK_START") or None

# TEST_MODE = True routes every email to TEST_EMAIL regardless of whose data
# it is, so you can check formatting safely. Flip to False to go live.
# Can also be set via the TEST_MODE env var ("true"/"false") - used by the
# GitHub Actions workflow_dispatch input so you can choose per-run without
# editing this file. Defaults to True (safe) if not set.
TEST_MODE = os.environ.get("TEST_MODE", "true").strip().lower() == "true"
TEST_EMAIL = "erik@tuxedofarmco.com"  # TODO: replace with your own address

FROM_NAME = "Tuxedo Farm Co."
REPLY_DEADLINE_TEXT = "by tomorrow"  # shown in the email body

# ---------------------------------------------------------------------------
# SCRIPT LOGIC - shouldn't need to touch below this line
# ---------------------------------------------------------------------------


def get_pay_week_range():
    """
    Return (week_start, week_end) as date objects for the pay week to pull.

    Auto mode (MANUAL_WEEK_START = None): always resolves to the most recently
    COMPLETED Monday-Sunday week, regardless of what day this is run on. E.g.
    running on a Tuesday still pulls last Mon-Sun, not this week's Mon/Tue.

    Manual mode: set MANUAL_WEEK_START to a specific Monday to pull that week
    instead (useful for re-sending or handling a late correction).
    """
    if MANUAL_WEEK_START:
        week_start = datetime.strptime(MANUAL_WEEK_START, "%Y-%m-%d").date()
    else:
        today = date.today()
        this_week_monday = today - timedelta(days=today.weekday())  # Monday = 0
        week_start = this_week_monday - timedelta(days=7)
    week_end = week_start + timedelta(days=6)  # Sunday
    return week_start, week_end


def parse_row_date(raw_value):
    """Parse a Smartsheet date cell value into a date object. Returns None if unparseable."""
    if not raw_value:
        return None
    if isinstance(raw_value, date):
        return raw_value
    text = str(raw_value).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            continue
    return None



def get_sheet_rows():
    """Pull all rows from the helper sheet as a list of {column_title: value} dicts."""
    ss = smartsheet.Smartsheet(SMARTSHEET_API_TOKEN)
    sheet = ss.Sheets.get_sheet(HELPER_SHEET_ID)
    col_id_to_title = {col.id: col.title for col in sheet.columns}

    rows_data = []
    for row in sheet.rows:
        row_dict = {}
        for cell in row.cells:
            title = col_id_to_title.get(cell.column_id)
            if title is None:
                continue
            value = cell.display_value if cell.display_value is not None else cell.value
            row_dict[title] = value
        rows_data.append(row_dict)
    return rows_data


def group_by_employee(rows, week_start, week_end):
    """Group rows into {(name, email): [row, row, ...]}, keeping only rows whose
    Date falls within [week_start, week_end] inclusive."""
    grouped = defaultdict(list)
    for row in rows:
        row_date = parse_row_date(row.get(COL_DATE))
        if row_date is None or not (week_start <= row_date <= week_end):
            continue

        name = row.get(COL_EMPLOYEE_NAME)
        email = row.get(COL_EMPLOYEE_EMAIL)
        if not name or not email:
            continue  # skip incomplete rows rather than guessing

        grouped[(name, email)].append(row)
    return grouped


def build_email_html(name, entries):
    """Build the HTML table + total for one employee's rows. Returns (html, total_hours)."""
    entries_sorted = sorted(entries, key=lambda r: str(r.get(COL_DATE) or ""))

    total_hours = 0.0
    table_rows = ""
    for entry in entries_sorted:
        entry_date = entry.get(COL_DATE, "") or ""
        job = entry.get(COL_JOB, "") or ""
        notes = entry.get(COL_NOTES, "") or ""
        raw_hours = entry.get(COL_HOURS, 0)
        try:
            hours = float(raw_hours)
        except (TypeError, ValueError):
            hours = 0.0
        total_hours += hours

        table_rows += f"""
        <tr>
            <td style="padding:6px 12px;border:1px solid #ddd;">{entry_date}</td>
            <td style="padding:6px 12px;border:1px solid #ddd;">{job}</td>
            <td style="padding:6px 12px;border:1px solid #ddd;text-align:right;">{hours:.2f}</td>
            <td style="padding:6px 12px;border:1px solid #ddd;">{notes}</td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
        <p>Hi {name},</p>
        <p>Here are your hours for the week. If anything looks wrong, please reply
           to this email {REPLY_DEADLINE_TEXT} — otherwise we'll treat this as
           confirmed and move ahead with payroll.</p>
        <table style="border-collapse:collapse;">
            <tr style="background:#f4f4f4;">
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Date</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Job/Field</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:right;">Hours</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Notes</th>
            </tr>
            {table_rows}
            <tr style="font-weight:bold;background:#f9f9f9;">
                <td style="padding:6px 12px;border:1px solid #ddd;" colspan="2">Total</td>
                <td style="padding:6px 12px;border:1px solid #ddd;text-align:right;">{total_hours:.2f}</td>
                <td style="padding:6px 12px;border:1px solid #ddd;"></td>
            </tr>
        </table>
        <p>Thanks,<br>{FROM_NAME}</p>
    </body>
    </html>
    """
    return html, total_hours


def send_email(to_address, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{FROM_NAME} <{GMAIL_ADDRESS}>"
    msg["To"] = to_address
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())


def main():
    week_start, week_end = get_pay_week_range()
    print(f"Pulling pay week: {week_start.isoformat()} to {week_end.isoformat()}")

    rows = get_sheet_rows()
    grouped = group_by_employee(rows, week_start, week_end)

    if not grouped:
        print("No matching rows found for that week — check the date range above "
              "against your sheet, and confirm COL_DATE matches your column title.")
        return

    print(f"Found {len(grouped)} employee(s) to send to. TEST_MODE = {TEST_MODE}")

    for (name, email), entries in grouped.items():
        html, total = build_email_html(name, entries)
        recipient = TEST_EMAIL if TEST_MODE else email
        subject = f"Weekly Hours Confirmation - {name} ({total:.2f} hrs)"
        send_email(recipient, subject, html)
        print(f"  Sent to {recipient} for {name}: {total:.2f} hours across {len(entries)} entries")


if __name__ == "__main__":
    main()
