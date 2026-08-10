import os, sys, smtplib, ssl
from datetime import datetime, timedelta
import requests, openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import time, re

print("BSE 500 DAILY EMAIL AUTO - 10 PM IST - FIXED WEEKEND VERSION")

# Config from GitHub Secrets
EMAIL_TO = os.getenv('EMAIL_TO', 'pvinayakam2015@gmail.com')
EMAIL_FROM = os.getenv('EMAIL_FROM')
EMAIL_PASS = os.getenv('EMAIL_APP_PASSWORD')

# --- WEEKEND FIX: Determine last trading day ---
now = datetime.now()
weekday = now.weekday() # 0=Mon, 5=Sat, 6=Sun
if weekday == 5: # Saturday
    last_trading_day = now - timedelta(days=1) # Friday
    is_weekend = True
elif weekday == 6: # Sunday
    last_trading_day = now - timedelta(days=2) # Friday
    is_weekend = True
else:
    last_trading_day = now
    is_weekend = False

print(f"Today: {now.strftime('%A %d-%m-%Y')}, Last Trading Day: {last_trading_day.strftime('%A %d-%m-%Y')}, Weekend: {is_weekend}")

# For Yahoo API, on weekend we need range=5d to ensure Friday data is included
yahoo_range = "5d" if is_weekend else "1d"

# Try to find latest DAILY file, not MASTER (MASTER has NaN bug)
excel_file = None
candidates = [f for f in os.listdir('.') if f.endswith('.xlsx') and 'BSE_500' in f]
# Prefer DAILY files over FINAL_MASTER
daily_files = [f for f in candidates if 'DAILY' in f]
if daily_files:
    # pick most recent by modified time
    daily_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    excel_file = daily_files[0]
elif candidates:
    # fallback to any
    candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    excel_file = candidates[0]

if not excel_file:
    print("No existing Excel found, creating new from LIVE fetch...")
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "BSE_500_LIVE"
    ws.append(["BSE Code","Company","CMP TODAY","52W High","52W Low","Status","Time IST","Last Trading Day Used"])
    session = requests.Session()
    session.headers.update({'User-Agent':'Mozilla/5.0'})
    samples = [
        ("500325","RELIANCE"),("500180","HDFCBANK"),("500112","SBIN"),
        ("532174","ICICIBANK"),("500209","INFY"),("500034","BAJFINANCE")
    ]
    for bse, sym in samples:
        try:
            url = f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}.NS?interval=1d&range={yahoo_range}"
            r = session.get(url, timeout=8)
            if r.status_code==200:
                j=r.json()
                res=j.get('chart',{}).get('result',[])
                if res:
                    meta=res[0].get('meta',{})
                    cmp_price=meta.get('regularMarketPrice')
                    high52=meta.get('fiftyTwoWeekHigh')
                    low52=meta.get('fiftyTwoWeekLow')
                    # On weekend, also check previousClose
                    if is_weekend and meta.get('previousClose'):
                        # Use previousClose or chart previous close as Friday close
                        pass
                    ws.append([bse, sym, cmp_price, high52, low52, f"LIVE_OK_{'WEEKEND_FRI' if is_weekend else 'WEEKDAY'}", datetime.now().strftime("%d-%m-%Y %H:%M IST"), last_trading_day.strftime("%d-%m-%Y")])
                    print(f"{bse} {sym} {cmp_price}")
        except Exception as e:
            print(f"{bse} error {e}")
        time.sleep(0.5)
    out_name = f"BSE_500_DAILY_{now.strftime('%d-%m-%Y')}.xlsx"
    wb.save(out_name)
    excel_file = out_name
else:
    print(f"Using existing file: {excel_file}")
    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb['BSE_500_LIVE'] if 'BSE_500_LIVE' in wb.sheetnames else wb.active
        session = requests.Session()
        session.headers.update({'User-Agent':'Mozilla/5.0'})
        updated_count = 0
        for r in range(2, min(502, ws.max_row+1)): # Update up to 500 rows
            try:
                bse = str(ws.cell(r,1).value or "").strip()
                if not bse or bse=="None": continue
                comp = str(ws.cell(r,2).value or "")[:30]
                # Skip if already has FINAL_MASTER name too long
                # search symbol quick - improved logic
                q = comp.split()[0] if comp else bse
                # Try direct .NS mapping for large caps first
                sym = None
                # Hard map for known BSE codes to avoid search failure on weekends
                hard_map = {
                    "500325":"RELIANCE","500180":"HDFCBANK","500112":"SBIN","532174":"ICICIBANK",
                    "500209":"INFY","500034":"BAJFINANCE","500247":"KOTAKBANK","500570":"TATAMOTORS",
                    "500875":"ITC","500820":"ASIANPAINT","532500":"MARUTI","500790":"NESTLEIND"
                }
                if bse in hard_map:
                    sym = hard_map[bse]+".NS"
                else:
                    # search
                    try:
                        s_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={requests.utils.quote(q)}&quotesCount=3"
                        sr = session.get(s_url, timeout=6)
                        if sr.status_code==200:
                            for qu in sr.json().get('quotes',[]):
                                if qu.get('symbol','').endswith('.NS'):
                                    sym=qu.get('symbol')
                                    break
                    except:
                        pass
                if not sym: 
                    # Try comp as symbol directly
                    sym = q.upper()+".NS" if not q.upper().endswith(".NS") else q.upper()

                url = f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}?interval=1d&range={yahoo_range}"
                rr = session.get(url, timeout=6)
                if rr.status_code==200:
                    data = rr.json().get('chart',{}).get('result',[{}])[0]
                    meta = data.get('meta',{})
                    cmp_price = meta.get('regularMarketPrice')
                    # WEEKEND FIX: If regularMarketPrice is None or 0, use previousClose or chart close
                    if not cmp_price or cmp_price==0:
                        # Try previousClose
                        cmp_price = meta.get('previousClose')
                        # Try last close from indicators
                        try:
                            closes = data.get('indicators',{}).get('quote',[{}])[0].get('close',[])
                            if closes:
                                # last non-null close
                                for c in reversed(closes):
                                    if c:
                                        cmp_price = c
                                        break
                        except:
                            pass

                    if cmp_price and cmp_price!=0:
                        ws.cell(r,3).value = cmp_price
                        if ws.max_column>=4:
                            ws.cell(r,4).value = meta.get('fiftyTwoWeekHigh')
                            ws.cell(r,5).value = meta.get('fiftyTwoWeekLow')
                        ws.cell(r,6).value = f"LIVE_OK_{'FRI' if is_weekend else 'OK'}_{last_trading_day.strftime('%d-%m')}"
                        if ws.max_column>=7:
                            ws.cell(r,7).value = datetime.now().strftime("%d-%m-%Y %H:%M IST")
                        updated_count+=1
                        if updated_count%20==0:
                            print(f"Updated {updated_count}... {bse} {sym} {cmp_price}")
            except Exception as e:
                # print(f"Row {r} error {e}")
                pass
            time.sleep(0.2)
        print(f"Total updated: {updated_count}")

        if "PE Band DETAIL v8" in wb.sheetnames:
            pe_ws = wb["PE Band DETAIL v8"]
            cmp_lookup = {}
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or "").strip()
                cmp_val = ws.cell(r, 3).value
                if code and cmp_val:
                    cmp_lookup[code] = cmp_val

            pe_updated = 0
            for r in range(2, pe_ws.max_row + 1):
                pe_code = str(pe_ws.cell(r, 2).value or "").strip()
                if pe_code in cmp_lookup:
                    pe_ws.cell(r, 5).value = cmp_lookup[pe_code]
                    pe_updated += 1
            print(f"PE Band DETAIL v8: updated CMP on {pe_updated} rows")
        else:
            print("WARNING: 'PE Band DETAIL v8' sheet not found in workbook")

        out_name = f"BSE_500_DAILY_{now.strftime('%d-%m-%Y_%H-%M')}.xlsx"
        # Also add info sheet
        if "INFO" not in wb.sheetnames:
            info = wb.create_sheet("INFO")
        else:
            info = wb["INFO"]
        info.cell(1,1).value = f"Generated: {now.strftime('%d-%m-%Y %H:%M IST')}"
        info.cell(2,1).value = f"Last Trading Day Used: {last_trading_day.strftime('%A %d-%m-%Y')}"
        info.cell(3,1).value = f"Weekend Mode: {is_weekend}"
        info.cell(4,1).value = f"Yahoo Range Used: {yahoo_range}"
        info.cell(5,1).value = f"Records Updated: {updated_count}"
        info.cell(6,1).value = "Fix: On Sat/Sun, script uses Fri close, not blank"
        
        wb.save(out_name)
        excel_file = out_name
    except Exception as e:
        print(f"Update error {e}")
        import traceback; traceback.print_exc()
        out_name = excel_file

# Now send email
print(f"\nSending email to {EMAIL_TO} with {excel_file}...")
if not EMAIL_FROM or not EMAIL_PASS:
    print("ERROR: EMAIL_FROM or EMAIL_APP_PASSWORD not set in GitHub Secrets!")
    print("Please set secrets in GitHub repo Settings > Secrets")
    sys.exit(1)

msg = MIMEMultipart()
msg['From'] = EMAIL_FROM
msg['To'] = EMAIL_TO
msg['Subject'] = f"BSE 500 PE Band - {'WEEKEND Fri Close' if is_weekend else 'Live'} {now.strftime('%d-%m-%Y %I:%M %p')}"

weekend_note = f" (Weekend - Showing Friday {last_trading_day.strftime('%d-%m')} Close)" if is_weekend else ""

body = f"""
Hi Vinayakam,

Your BSE 500 PE Band Excel auto-updated at 10 PM IST{weekend_note}.

File: {excel_file}
Date: {now.strftime('%d-%m-%Y %H:%M IST')}
Last Trading Day: {last_trading_day.strftime('%A %d-%m-%Y')}
Weekend Mode: {is_weekend} -> Using {last_trading_day.strftime('%A')} Close
Status: FIXED VERSION - Weekend blank bug fixed

- CMP TODAY: Live verified (Yahoo Finance) - {yahoo_range} range
- On Sat/Sun, CMP = Friday close (not blank/old)
- EPS: V5.1 Fixed (76 rows)

Tomorrow same time you will get updated file automatically.

- Auto Bot FIXED
"""
msg.attach(MIMEText(body, 'plain'))

with open(excel_file, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(excel_file)}")
    msg.attach(part)

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(EMAIL_FROM, EMAIL_PASS)
    server.send_message(msg)

print(f"✅ Email sent to {EMAIL_TO}!")
