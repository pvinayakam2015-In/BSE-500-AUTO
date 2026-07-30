
import os, smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pandas as pd
from datetime import datetime

# Config from GitHub Secrets
EMAIL_FROM = os.environ.get("EMAIL_FROM", "pvinayakam2015@gmail.com")
EMAIL_TO = os.environ.get("EMAIL_TO", "pvinayakam2015@gmail.com")
EMAIL_APP_PASSWORD = os.environ.get("EMAIL_APP_PASSWORD", "")
EXCEL_FILE = "BSE_500_FINAL_MASTER_START.xlsx"

print(f"Starting daily update: {datetime.now()}")

# For now, just send existing file (price update logic already in file)
# If you have auto-update logic, it will run here
try:
    # Simple check file exists
    if not os.path.exists(EXCEL_FILE):
        # try alternative names
        for alt in ["BSE_500_FINAL_MASTER_V5_1_76EPS_FIXED.xlsx", "BSE_500_FINAL_MASTER_V5_1_76EPS_FIXED_30-07-2026.xlsx"]:
            if os.path.exists(alt):
                EXCEL_FILE = alt
                break
    print(f"Using file: {EXCEL_FILE}")

    # Create email
    msg = MIMEMultipart()
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg["Subject"] = f"BSE 500 Auto Update - {datetime.now().strftime('%d-%m-%Y %I:%M %p')}"

    body = f"""Hi,

Your BSE 500 auto-updated Excel is ready.

File: {EXCEL_FILE}
Date: {datetime.now().strftime('%d-%m-%Y %I:%M %p IST')}

This email is auto-generated daily at 10 PM IST via GitHub Actions.

Regards,
BSE 500 AUTO System
"""
    msg.attach(MIMEText(body, "plain"))

    # Attach Excel
    with open(EXCEL_FILE, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(EXCEL_FILE)}")
    msg.attach(part)

    # Send via Gmail SMTP
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(EMAIL_FROM, EMAIL_APP_PASSWORD)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
    
    print("Email sent successfully!")

except Exception as e:
    print(f"ERROR: {e}")
    raise
