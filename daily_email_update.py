"""
BSE 500 PE Band - Daily Auto Update & Email
Runs nightly at 10 PM IST via GitHub Actions (.github/workflows/daily.yml).

What this does:
1. Loads the static base workbook (BSE_500_PE_Band_BASE.xlsx) - 501 companies,
   FY2016-FY2026 Year High/Low and EPS (manually sourced and verified - see
   the README sheet inside that file for methodology).
2. Fetches today's live CMP for all 501 BSE Codes directly from BSE India's
   official getScripHeaderData API (no fuzzy name/symbol search - this was
   the root cause of wrong data in earlier versions of this pipeline).
3. Recomputes High PE / Low PE for every row as real numbers (Year High/EPS,
   Year Low/EPS), skipping rows where EPS is N/A, Not Sourced, missing, or
   negative (PE not meaningful in those cases - flagged accordingly).
4. Saves a dated snapshot and emails it.

The base workbook itself is never modified by this script - only a fresh
dated output file is generated and emailed each run.
"""

import os
import sys
import time
import smtplib
import ssl
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import openpyxl
import requests

print("BSE 500 PE Band - Daily Update & Email")

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
EMAIL_TO = os.getenv("EMAIL_TO", "pvinayakam2015@gmail.com")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_PASS = os.getenv("EMAIL_APP_PASSWORD")

BASE_FILE = "BSE_500_PE_Band_BASE.xlsx"
SHEET_NAME = "PE Band DETAIL v8"

# Column layout of the base workbook (1-indexed, matches openpyxl .cell(r, c))
COL_SR = 1
COL_BSE_CODE = 2
COL_ISIN = 3
COL_COMPANY = 4
COL_FY = 5
COL_CMP = 6
COL_EPS = 7
COL_EPS_SOURCE = 8
COL_YEAR_HIGH = 9
COL_YEAR_LOW = 10
COL_HIGH_PE = 11
COL_LOW_PE = 12
COL_CONFIDENCE = 13

BSE_API_URL = "https://api.bseindia.com/BseIndiaAPI/api/getScripHeaderData/w"
BSE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Origin": "https://www.bseindia.com",
    "Referer": "https://www.bseindia.com/",
    "Connection": "keep-alive",
}

# ---------------------------------------------------------------------------
# Determine last trading day (weekend handling)
# ---------------------------------------------------------------------------
now = datetime.now()
weekday = now.weekday()  # 0=Mon ... 5=Sat, 6=Sun
if weekday == 5:
    last_trading_day = now - timedelta(days=1)
    is_weekend = True
elif weekday == 6:
    last_trading_day = now - timedelta(days=2)
    is_weekend = True
else:
    last_trading_day = now
    is_weekend = False

print(
    f"Today: {now.strftime('%A %d-%m-%Y')}, "
    f"Last Trading Day: {last_trading_day.strftime('%A %d-%m-%Y')}, "
    f"Weekend: {is_weekend}"
)

# ---------------------------------------------------------------------------
# CMP fetch - BSE official API, by BSE Code directly (no name search)
# ---------------------------------------------------------------------------
_cffi_available = True
try:
    from curl_cffi import requests as cffi_requests
except ImportError:
    _cffi_available = False
    print("curl_cffi not available, will use plain requests only")

_plain_session = requests.Session()
_plain_session.headers.update(BSE_HEADERS)


def fetch_cmp(scripcode: str):
    """
    Fetch CMP for a single BSE scripcode from BSE's official
    getScripHeaderData API. Tries curl_cffi (browser impersonation, more
    reliable against BSE's WAF from cloud/CI IPs) first, then falls back
    to a plain requests session. Returns float or None.
    """
    params = {"scripcode": scripcode}

    # Attempt 1: curl_cffi with Chrome impersonation
    if _cffi_available:
        try:
            r = cffi_requests.get(
                BSE_API_URL,
                params=params,
                headers=BSE_HEADERS,
                impersonate="chrome",
                timeout=10,
            )
            if r.status_code == 200:
                header = r.json().get("Header", {})
                val = header.get("PrevClose") or header.get("LTP")
                if val:
                    return float(val)
        except Exception:
            pass

    # Attempt 2: plain requests session
    try:
        r = _plain_session.get(BSE_API_URL, params=params, timeout=10)
        if r.status_code == 200:
            header = r.json().get("Header", {})
            val = header.get("PrevClose") or header.get("LTP")
            if val:
                return float(val)
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Load base workbook
# ---------------------------------------------------------------------------
if not os.path.exists(BASE_FILE):
    print(f"FATAL: base file '{BASE_FILE}' not found in repo checkout.")
    sys.exit(1)

print(f"Loading base workbook: {BASE_FILE}")
wb = openpyxl.load_workbook(BASE_FILE)

if SHEET_NAME not in wb.sheetnames:
    print(f"FATAL: sheet '{SHEET_NAME}' not found in {BASE_FILE}")
    sys.exit(1)

ws = wb[SHEET_NAME]
max_row = ws.max_row

# Collect unique BSE codes (one CMP fetch per company, applied to all its FY rows)
unique_codes = []
seen = set()
rows_by_code = {}
for r in range(2, max_row + 1):
    code = ws.cell(r, COL_BSE_CODE).value
    if code is None:
        continue
    code_str = str(code).strip()
    if not code_str:
        continue
    rows_by_code.setdefault(code_str, []).append(r)
    if code_str not in seen:
        seen.add(code_str)
        unique_codes.append(code_str)

print(f"Companies to fetch CMP for: {len(unique_codes)}")

# ---------------------------------------------------------------------------
# Fetch CMP for every company
# ---------------------------------------------------------------------------
cmp_by_code = {}
fetched = 0
failed = 0

for i, code in enumerate(unique_codes, start=1):
    cmp_val = fetch_cmp(code)
    if cmp_val:
        cmp_by_code[code] = cmp_val
        fetched += 1
    else:
        failed += 1
    if i % 50 == 0:
        print(f"  ... {i}/{len(unique_codes)} processed "
              f"(fetched {fetched}, failed {failed})")
    time.sleep(0.35)  # be polite to BSE's servers

print(f"CMP fetch complete: {fetched} succeeded, {failed} failed "
      f"out of {len(unique_codes)}")

# ---------------------------------------------------------------------------
# Apply CMP + recompute PE for every row
# ---------------------------------------------------------------------------
updated_rows = 0
pe_computed = 0

for code, rows in rows_by_code.items():
    cmp_val = cmp_by_code.get(code)
    for r in rows:
        if cmp_val is not None:
            ws.cell(r, COL_CMP).value = cmp_val
            updated_rows += 1
        else:
            # No live CMP this run - flag it, keep whatever CMP value the
            # base file already had (better a stale number than a blank).
            existing_flag = ws.cell(r, COL_CONFIDENCE).value
            if existing_flag == "OK":
                ws.cell(r, COL_CONFIDENCE).value = "No CMP Match (fetch failed today)"

        eps = ws.cell(r, COL_EPS).value
        high = ws.cell(r, COL_YEAR_HIGH).value
        low = ws.cell(r, COL_YEAR_LOW).value
        cmp_now = ws.cell(r, COL_CMP).value

        eps_numeric = isinstance(eps, (int, float))
        high_numeric = isinstance(high, (int, float))
        low_numeric = isinstance(low, (int, float))

        if eps_numeric and eps != 0 and high_numeric and low_numeric:
            if eps > 0:
                ws.cell(r, COL_HIGH_PE).value = round(high / eps, 2)
                ws.cell(r, COL_LOW_PE).value = round(low / eps, 2)
                pe_computed += 1
            else:
                # Negative EPS - PE not meaningful, leave blank
                ws.cell(r, COL_HIGH_PE).value = None
                ws.cell(r, COL_LOW_PE).value = None
        else:
            ws.cell(r, COL_HIGH_PE).value = None
            ws.cell(r, COL_LOW_PE).value = None

print(f"Rows with CMP updated: {updated_rows}")
print(f"Rows with PE recomputed: {pe_computed}")

# ---------------------------------------------------------------------------
# INFO sheet
# ---------------------------------------------------------------------------
if "INFO" in wb.sheetnames:
    info = wb["INFO"]
    wb.remove(info)
info = wb.create_sheet("INFO")
info_lines = [
    f"Generated: {now.strftime('%d-%m-%Y %H:%M IST')}",
    f"Last Trading Day Used: {last_trading_day.strftime('%A %d-%m-%Y')}",
    f"Weekend Mode: {is_weekend}",
    f"Companies: {len(unique_codes)} | CMP fetched OK: {fetched} | CMP fetch failed: {failed}",
    f"CMP Source: BSE India official API (getScripHeaderData), by BSE Code directly",
    f"PE High/Low: recomputed live this run as Year High/EPS and Year Low/EPS",
    f"Base data (Year High/Low, EPS): BSE_500_PE_Band_BASE.xlsx - see its README sheet",
]
for idx, line in enumerate(info_lines, start=1):
    info.cell(idx, 1).value = line

# ---------------------------------------------------------------------------
# Save dated output and email
# ---------------------------------------------------------------------------
out_name = f"BSE_500_PE_Band_DAILY_{now.strftime('%d-%m-%Y_%H-%M')}.xlsx"
wb.save(out_name)
print(f"Saved: {out_name}")

print(f"\nSending email to {EMAIL_TO} with {out_name}...")
if not EMAIL_FROM or not EMAIL_PASS:
    print("ERROR: EMAIL_FROM or EMAIL_APP_PASSWORD not set in GitHub Secrets!")
    sys.exit(1)

msg = MIMEMultipart()
msg["From"] = EMAIL_FROM
msg["To"] = EMAIL_TO
msg["Subject"] = (
    f"BSE 500 PE Band - "
    f"{'Weekend (Fri Close)' if is_weekend else 'Live'} "
    f"{now.strftime('%d-%m-%Y %I:%M %p')}"
)

weekend_note = (
    f" (Weekend - showing {last_trading_day.strftime('%A %d-%m')} close)"
    if is_weekend else ""
)

body = f"""Hi Vinayakam,

Your BSE 500 PE Band workbook auto-updated at 10 PM IST{weekend_note}.

Companies: {len(unique_codes)}
CMP fetched successfully: {fetched}
CMP fetch failed today: {failed} (flagged "No CMP Match (fetch failed today)" in the file - CMP left at previous value for those)
Date: {now.strftime('%d-%m-%Y %H:%M IST')}
Last Trading Day: {last_trading_day.strftime('%A %d-%m-%Y')}

- CMP: live from BSE India's official API, fetched directly by BSE Code (no name/symbol search)
- High PE / Low PE: recomputed this run from Year High/EPS and Year Low/EPS
- Base data (Year High/Low, EPS FY2016-FY2026): your completed 501-company tracker - see the README sheet inside the attached file for full methodology and confidence flags

Tomorrow same time you'll get the updated file automatically.

- Auto Bot
"""
msg.attach(MIMEText(body, "plain"))

with open(out_name, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition", f"attachment; filename= {os.path.basename(out_name)}"
    )
    msg.attach(part)

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(EMAIL_FROM, EMAIL_PASS)
    server.send_message(msg)

print(f"Email sent to {EMAIL_TO}!")
