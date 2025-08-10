# -*- coding: utf-8 -*-
import os
import json
from imapclient import IMAPClient
import pyzmail
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from zoneinfo import ZoneInfo

# ================= CONFIG =================
EMAIL = os.getenv("EMAIL_ACCOUNT")
PASSWORD = os.getenv("EMAIL_PASSWORD")
SPREADSHEET_URL = os.getenv("SPREADSHEET_URL")
SPREADSHEET_NAME = "E-money Tracker"
WORKSHEET_NAME = "Data"

# ================= SERVICE ACCOUNT JSON =================

SERVICE_ACCOUNT_STR = os.getenv("GOOGLE_CREDENTIALS")
SERVICE_ACCOUNT_DICT = json.loads(SERVICE_ACCOUNT_STR)
CREDENTIALS_JSON = Credentials.from_service_account_info(SERVICE_ACCOUNT_DICT)

# ================= FUNGSI EKSTRAKSI =================
def extract_nominal(text):
    match = re.search(r'Rp\s?[\d\.]+', text, re.IGNORECASE)
    if match:
        angka = match.group(0).replace(" ", "")
        return angka
    return None

def extract_ref_number(text):
    match = re.search(r'\b\d{18}\b', text)
    return match.group(0) if match else None

# ================= EMAIL SCRAPER =================
today_str = datetime.now(ZoneInfo("Asia/Jakarta")).strftime('%Y-%m-%d')  # ✅ WIB
results = []

with IMAPClient('imap.gmail.com') as client:
    client.login(EMAIL, PASSWORD)
    client.select_folder('INBOX')

    messages = client.search(['FROM', 'noreply.livin@bankmandiri.co.id'])
    raw_data = client.fetch(messages, ['BODY[]', 'ENVELOPE'])

    for uid in messages:
        envelope = raw_data[uid][b'ENVELOPE']
        subject = envelope.subject.decode() if envelope.subject else ""

        # ✅ Konversi tanggal ke WIB
        email_date_wib = envelope.date.astimezone(ZoneInfo("Asia/Jakarta"))

        # Hanya proses email hari ini
        if email_date_wib.strftime('%Y-%m-%d') != today_str:
            continue

        if re.search(r"top[-\s]?up", subject, re.IGNORECASE):
            msg = pyzmail.PyzMessage.factory(raw_data[uid][b'BODY[]'])

            if msg.text_part:
                body = msg.text_part.get_payload().decode(msg.text_part.charset)
            elif msg.html_part:
                body = msg.html_part.get_payload().decode(msg.html_part.charset)
            else:
                body = ""

            nominal = extract_nominal(body)
            ref_number = extract_ref_number(body)

            results.append([
                email_date_wib.strftime('%Y-%m-%d %H:%M:%S'),  # ✅ sudah WIB
                subject,
                nominal,
                ref_number
            ])

# ================= APPEND KE GOOGLE SHEETS =================
if results:
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = CREDENTIALS_JSON.with_scopes(scopes)
    client_gs = gspread.authorize(creds)

    sheet = client_gs.open_by_url(SPREADSHEET_URL).worksheet("Data")
    sheet.append_rows(results)

    print(f"{len(results)} data baru ditambahkan ke Google Sheets.")
else:
    print("Tidak ada data top-up hari ini.")