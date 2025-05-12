import pandas as pd
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os
from datetime import datetime
import time

# Load environment variables
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

EXCEL_FILE = "clients.xlsx"
HTML_TEMPLATE_FILE = "email_template.html"
EMAIL_COLUMN = "Email"

SUBJECT = "🌟 Special Offer for You!"

# Read email list from Excel
def read_emails(file_path):
    df = pd.read_excel(file_path)
    return df[EMAIL_COLUMN].dropna().unique().tolist()

# Load HTML template
def load_html_template():
    with open(HTML_TEMPLATE_FILE, "r", encoding="utf-8") as file:
        return file.read()

# Send email
def send_email(to_email, html_content):
    msg = EmailMessage()
    msg["Subject"] = SUBJECT
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email

    msg.set_content("This is an HTML email. Please use an email client that supports HTML.")
    msg.add_alternative(html_content, subtype="html")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print(f"[{datetime.now()}] ✅ Sent to: {to_email}")
    except Exception as e:
        print(f"[{datetime.now()}] ❌ Failed to send to {to_email}: {e}")

# Main execution
if __name__ == "__main__":
    recipients = read_emails(EXCEL_FILE)
    html_template = load_html_template()

    print(f"📬 Sending to {len(recipients)} recipients...\n")

    for email in recipients:
        send_email(email, html_template)
        time.sleep(2)  # avoid spam flag
