import requests
import pandas as pd
import datetime
import win32com.client
import html  # Needed for html.escape()
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
# === CONFIGURATION ===

BUGZILLA_API_URL=os.getenv("BUGZILLA_API_URL")
BUGZILLA_API_KEY=os.getenv("BUGZILLA_API_KEY")
ASSIGNED_TO_EMAIL = ""  # Or leave blank for all
PRODUCT_NAME =os.getenv("SEZ_PRODUCT_NAME")
BUG_STATUSES =os.getenv("SEZ_BUG_STATUSES").split(",")
TO_EMAILS =os.getenv("SEZ_TO_EMAILS").split(',')
CC_EMAILS =os.getenv("SEZ_CC_EMAILS").split(',')

# === FETCH BUGS FROM BUGZILLA ===

def fetch_open_bugs():
    params = {
        "api_key": BUGZILLA_API_KEY,
        "status": BUG_STATUSES,
        "include_fields": "id,summary,priority,assigned_to,component,cf_planned_release,status,last_change_time,assigned_to_detail"
    }

    if ASSIGNED_TO_EMAIL:
        params["assigned_to"] = ASSIGNED_TO_EMAIL
    if PRODUCT_NAME:
        params["product"] = PRODUCT_NAME
    print(params)
    response = requests.get(f"{BUGZILLA_API_URL}/bug", params=params)
    response.raise_for_status()
    return response.json().get("bugs", [])

# === GENERATE HTML REPORT ===

def generate_html_report(bugs):
    if not bugs:
        return "<p>No open bugs found.</p>"

    cleaned_bugs = []
    for bug in bugs:
        assigned_to_detail = bug.get("assigned_to_detail", {})
        cleaned_bug = {
            "ID": bug.get("id"),
            "Summary": html.escape(bug.get("summary", "")),
            "Priority": bug.get("priority", ""),
            "Assigned To": assigned_to_detail.get("real_name") or bug.get("assigned_to", ""),
            "Component": bug.get("component", ""),
            "Release": bug.get("cf_planned_release", ""),
            "Status": bug.get("status", ""),
            "Last Updated": bug.get("last_change_time", "")[:10],
        }
        # if len(cleaned_bug["Summary"]) > 100:
        #     cleaned_bug["Summary"] = cleaned_bug["Summary"][:97] + "..."
        cleaned_bugs.append(cleaned_bug)

    df = pd.DataFrame(cleaned_bugs).sort_values(by="ID")

    # Summary: Bugs per user
    summary_table = df["Assigned To"].value_counts().reset_index()
    summary_table.columns = ["Assigned To", "Bug Count"]
    summary_html = "<h3>Summary: Bugs Assigned per User</h3><table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; font-family: Arial;'>"
    summary_html += "<tr style='background-color: #1a73e8; color: white;'><th>Assigned To</th><th>Bug Count</th></tr>"
    for _, row in summary_table.iterrows():
        summary_html += f"<tr><td>{row['Assigned To']}</td><td>{row['Bug Count']}</td></tr>"
    summary_html += "</table><br>"

    # Main Bug Table
    columns = ["ID", "Summary", "Priority", "Assigned To", "Component", "Release", "Status", "Last Updated"]
    column_styles = {
        "ID": "min-width:60px;",
        "Summary": "min-width:300px;",
        "Priority": "min-width:80px;",
        "Assigned To": "min-width:160px;",
        "Component": "min-width:120px;",
        "Release": "min-width:80px;",
        "Status": "min-width:150px;",  # FIXED WIDTH FOR LONG STATUSES
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
            value = html.escape(str(row[col]))
            html_table += f"<td style='{column_styles.get(col, '')} white-space: normal;'>{value}</td>"
        html_table += "</tr>"
    html_table += "</table>"

    html_report = f"""
    <html>
    <head>
    <meta charset="UTF-8">
    </head>
    <body>
        <h2>Bugzilla Bug Report - {datetime.date.today()}</h2>
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
        to_emails=TO_EMAILS,
        cc_emails=CC_EMAILS
    )

if __name__ == "__main__":
    main()