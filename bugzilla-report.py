import requests
import pandas as pd
import datetime
import win32com.client
import html  # Needed for html.escape()
from datetime import timedelta, timezone
import pytz
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

BUGZILLA_API_URL = os.getenv("BUGZILLA_API_URL")
BUGZILLA_API_KEY = os.getenv("BUGZILLA_API_KEY")
PRODUCT_NAME = os.getenv("PRODUCT_NAME")
BUG_STATUSES = os.getenv("BUG_STATUSES").split(",")
TO_EMAILS = os.getenv("TO_EMAILS").split(",")
CC_EMAILS = os.getenv("CC_EMAILS").split(",")
ASSIGNED_TO_EMAIL=""
print("URL:", BUGZILLA_API_URL)  

# === FETCH BUGS FROM BUGZILLA ===
def fetch_open_bugs():
    params = {
        "api_key": BUGZILLA_API_KEY,
        "status": [status.strip() for status in BUG_STATUSES],
        "include_fields": "id,summary,priority,assigned_to,component,cf_planned_release,status,cf_target_date,assigned_to_detail,last_change_time"
    }

    if ASSIGNED_TO_EMAIL:
        params["assigned_to"] = ASSIGNED_TO_EMAIL
    if PRODUCT_NAME:
        params["product"] = PRODUCT_NAME

    response = requests.get(f"{BUGZILLA_API_URL}/bug", params=params)
    response.raise_for_status()
    return response.json().get("bugs", [])


def generate_html_report(bugs):
    if not bugs:
        return "<p>No open bugs found.</p>"

    # Define IST timezone
    ist_tz = pytz.timezone("Asia/Kolkata")
    today = datetime.datetime.now(ist_tz).date()

    def convert_to_ist_date(date_str):
        """Convert date string from Bugzilla UTC to IST date."""
        if not date_str:
            return None
        try:
            # Parse as UTC
            dt_utc = datetime.datetime.strptime(date_str[:19], "%Y-%m-%dT%H:%M:%S")
            dt_utc = dt_utc.replace(tzinfo=pytz.UTC)
            # Convert to IST
            dt_ist = dt_utc.astimezone(ist_tz)
            return dt_ist.date()
        except Exception:
            return None

    cleaned_bugs = []
    for bug in bugs:
        assigned_to_detail = bug.get("assigned_to_detail", {})

        target_date_obj = convert_to_ist_date(bug.get("cf_target_date"))
        last_updated_obj = convert_to_ist_date(bug.get("last_change_time"))

        cleaned_bug = {
            "ID": bug.get("id"),
            "Summary": html.escape(bug.get("summary", "")),
            "Priority": bug.get("priority", ""),
            "Assigned To": assigned_to_detail.get("real_name") or bug.get("assigned_to", ""),
            "Component": bug.get("component", ""),
            "Release": bug.get("cf_planned_release", ""),
            "Status": bug.get("status", ""),
            "Target Date": target_date_obj,
            "Last Updated": last_updated_obj.strftime("%d/%m/%Y") if last_updated_obj else ""
        }
        cleaned_bugs.append(cleaned_bug)

    df = pd.DataFrame(cleaned_bugs).sort_values(by="ID")

    # Summary table
    summary_table = df["Assigned To"].value_counts().reset_index()
    summary_table.columns = ["Assigned To", "Bug Count"]
    summary_html = "<h3>Summary: Bugs Assigned per User</h3><table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; font-family: Arial;'>"
    summary_html += "<tr style='background-color: #1a73e8; color: white;'><th>Assigned To</th><th>Bug Count</th></tr>"
    for _, row in summary_table.iterrows():
        summary_html += f"<tr><td>{row['Assigned To']}</td><td>{row['Bug Count']}</td></tr>"
    summary_html += "</table><br>"

    # Table structure
    columns = ["ID", "Summary", "Priority", "Assigned To", "Component", "Release", "Status", "Target Date", "Last Updated"]
    column_styles = {
        "ID": "min-width:60px;",
        "Summary": "min-width:300px;",
        "Priority": "min-width:80px;",
        "Assigned To": "min-width:160px;",
        "Component": "min-width:120px;",
        "Release": "min-width:80px;",
        "Status": "min-width:150px;",
        "Target Date": "min-width:110px;",
        "Last Updated": "min-width:100px;"
    }

    html_table = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; font-family: Arial; width: 100%; table-layout: auto;'>"
    html_table += "<tr style='background-color: #1a73e8; color: white;'>"
    for col in columns:
        html_table += f"<th style='{column_styles.get(col, '')} text-align: left;'>{col}</th>"
    html_table += "</tr>"

    for _, row in df.iterrows():
        html_table += "<tr>"
        for col in columns:
            if col == "Target Date":
                target_display = ""
                cell_style = column_styles.get(col, "") + " white-space: normal;"

                if isinstance(row["Target Date"], datetime.date):
                    target_display = row["Target Date"].strftime("%d/%m/%Y")
                    if row["Target Date"] < today:
                        cell_style += " background-color: red; color: white;"
                    elif row["Target Date"] == today:
                        cell_style += " background-color: yellow; color: black;"
                    else:
                        cell_style += " background-color: orange; color: black;"

                html_table += f"<td style='{cell_style}'>{target_display}</td>"
            else:
                val = row[col] if row[col] is not None else ""
                html_table += f"<td style='{column_styles.get(col, '')} white-space: normal;'>{html.escape(str(val))}</td>"
        html_table += "</tr>"

    html_table += "</table>"

    html_report = f"""
    <html>
    <head>
    <meta charset="UTF-8">
    </head>
    <body>
        <h2>Bugzilla Bug Report - {today.strftime('%d/%m/%Y')}</h2>
        <p>Total Open Bugs: <b>{len(df)}</b></p>
        {summary_html}
        {html_table}
        <p>--<br>This report was auto-generated.</p>
    </body>
    </html>
    """
    return html_report


# === SEND EMAIL VIA OUTLOOK ===

def send_email_via_outlook(html_body, subject, to_emails, cc_emails):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(to_emails)
    mail.CC = "; ".join(cc_emails)
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Display()  # Keeps it interactive; change to mail.Send() for auto-send
    print("Outlook draft created with TO and CC.")

# === MAIN ===

def main():
    bugs = fetch_open_bugs()
    html_report = generate_html_report(bugs)
    send_email_via_outlook(
        html_body=html_report,
        subject=f"{PRODUCT_NAME} Daily Bugzilla Bug Report - {datetime.date.today()}",
        to_emails=[email.strip() for email in TO_EMAILS],
        cc_emails=[email.strip() for email in CC_EMAILS]
    )

if __name__ == "__main__":
    main()