"""
Weekly Hours Confirmation Emailer
----------------------------------
Pulls time entries from a Smartsheet helper sheet, groups them by employee,
builds a clean HTML table (Date / Job / Hours / Notes + weekly Total), and
emails each employee their week's hours for confirmation before payroll.

Sends via the Gmail API using a service account with domain-wide delegation
(Google Workspace no longer supports App Passwords / SMTP as of March 2025).

Setup:
    pip install smartsheet-python-sdk google-auth google-api-python-client

Required environment variables (set as secrets if running in GitHub Actions):
    SMARTSHEET_API_TOKEN        - your Smartsheet API access token
    GMAIL_ADDRESS                - the Workspace address the emails are sent from
    GOOGLE_SERVICE_ACCOUNT_JSON  - full contents of the service account JSON key
                                   (Cloud Console > IAM & Admin > Service Accounts
                                   > Keys > Add Key > JSON), authorized for domain-wide
                                   delegation with scope gmail.send in the Workspace
                                   Admin console (Security > API Controls > Domain-wide
                                   Delegation)
"""

import base64
import json
import os
import smtplib  # noqa: F401 (kept for reference; no longer used for sending)
from collections import defaultdict
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import smartsheet
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# CONFIG - edit these to match your helper sheet
# ---------------------------------------------------------------------------

SMARTSHEET_API_TOKEN = os.environ["SMARTSHEET_API_TOKEN"]
GMAIL_ADDRESS = os.environ["GMAIL_ADDRESS"]  # the Workspace mailbox sending as
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

HELPER_SHEET_ID = 6881750327185284

# TODO: these must match your helper sheet's column titles exactly
COL_EMPLOYEE_NAME = "Name"
COL_EMPLOYEE_EMAIL = "Email"
COL_DATE = "Date"
COL_JOB = "Customer Job"
COL_PAYROLL_ITEM = "Payroll Item"
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
REPLY_DEADLINE_TEXT = "antes de mañana"  # shown in the email body

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
        payroll_item = entry.get(COL_PAYROLL_ITEM, "") or ""
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
            <td style="padding:6px 12px;border:1px solid #ddd;">{payroll_item}</td>
            <td style="padding:6px 12px;border:1px solid #ddd;text-align:right;">{hours:.2f}</td>
            <td style="padding:6px 12px;border:1px solid #ddd;">{notes}</td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:Arial,sans-serif;font-size:14px;color:#222;">
        <p>Hola {name},</p>
        <p>Aquí están sus horas de esta semana. Si algo no le parece correcto,
           por favor responda a este correo {REPLY_DEADLINE_TEXT} — de lo contrario
           entenderemos que está confirmado y seguiremos adelante con la nómina.</p>
        <table style="border-collapse:collapse;">
            <tr style="background:#f4f4f4;">
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Fecha</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Campo</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Trabajo</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:right;">Horas</th>
                <th style="padding:6px 12px;border:1px solid #ddd;text-align:left;">Notas</th>
            </tr>
            {table_rows}
            <tr style="font-weight:bold;background:#f9f9f9;">
                <td style="padding:6px 12px;border:1px solid #ddd;" colspan="3">Total</td>
                <td style="padding:6px 12px;border:1px solid #ddd;text-align:right;">{total_hours:.2f}</td>
                <td style="padding:6px 12px;border:1px solid #ddd;"></td>
            </tr>
        </table>
        <p>Gracias,<br>{FROM_NAME}</p>
    </body>
    </html>
    """
    return html, total_hours


def get_gmail_service():
    """Build an authorized Gmail API client, impersonating GMAIL_ADDRESS via
    domain-wide delegation (must be authorized in the Workspace Admin console)."""
    service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    credentials = service_account.Credentials.from_service_account_info(
        service_account_info, scopes=GMAIL_SCOPES
    ).with_subject(GMAIL_ADDRESS)
    return build("gmail", "v1", credentials=credentials, cache_discovery=False)


def send_email(gmail_service, to_address, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{FROM_NAME} <{GMAIL_ADDRESS}>"
    msg["To"] = to_address
    msg.attach(MIMEText(html_body, "html"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    gmail_service.users().messages().send(userId="me", body={"raw": raw}).execute()


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

    gmail_service = get_gmail_service()

    for (name, email), entries in grouped.items():
        html, total = build_email_html(name, entries)
        recipient = TEST_EMAIL if TEST_MODE else email
        subject = f"Confirmación de Horas Semanales - {name} ({total:.2f} hrs)"
        send_email(gmail_service, recipient, subject, html)
        print(f"  Sent to {recipient} for {name}: {total:.2f} hours across {len(entries)} entries")


if __name__ == "__main__":
    main()
