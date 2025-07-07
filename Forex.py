import requests
import pdfplumber
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import socket
from openpyxl import Workbook
from datetime import datetime

# ---------------- INTERNET CHECK ----------------
def check_internet(timeout=5):
    try:
        socket.setdefaulttimeout(timeout)
        host = socket.gethostbyname("www.google.com")
        s = socket.create_connection((host, 80), timeout)
        s.close()
        return True
    except Exception:
        return False

# ---------------- SBI FUNCTION ----------------
def fetch_sbi_all_slabs():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--window-size=1200,800")
    driver = webdriver.Chrome(options=options)

    results = []
    try:
        driver.get("https://sbinewyork.statebank/exchange-rate")
        time.sleep(7)

        tables = driver.find_elements(By.TAG_NAME, "table")
        for table in tables:
            headers = table.find_elements(By.TAG_NAME, "th")
            if any("Remittance Amount" in h.text for h in headers):
                rows = table.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 2:
                        slab = cells[0].text.strip()
                        try:
                            rate = float(cells[1].text.strip())
                            results.append(("SBI", slab, rate))
                        except ValueError:
                            continue
        driver.quit()
        return results
    except Exception as e:
        print("Error fetching SBI rates:", e)
        driver.quit()
        return []

# ---------------- HDFC FUNCTION ----------------
def fetch_hdfc_usd_cash_buying():
    pdf_url = "https://www.hdfcbank.com/content/bbp/repositories/723fb80a-2dde-42a3-9793-7ae1be57c87f/?path=%2FPersonal%2FHome%2Fcontent%2Frates.pdf"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    try:
        response = requests.get(pdf_url, headers=headers, timeout=15)
        with open("hdfc_rates.pdf", "wb") as f:
            f.write(response.content)

        with pdfplumber.open("hdfc_rates.pdf") as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                for line in text.split('\n'):
                    if "United States Dollar USD" in line:
                        parts = line.split()
                        for i in range(len(parts)):
                            if parts[i] == "USD":
                                try:
                                    rate = float(parts[i + 3])
                                    return [("HDFC", "Cash Buying", rate)]
                                except (ValueError, IndexError):
                                    return []
        return []
    except Exception as e:
        print("Error fetching HDFC rate:", e)
        return []

# ---------------- EXCEL OUTPUT ----------------
def save_rates_to_excel(data):
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"forex_rates_{date_str}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Forex Rates"
    ws.append(["Bank", "Slab / Type", "USD to INR Rate"])

    for row in data:
        ws.append(row)

    wb.save(filename)
    print(f"✅ Saved: {filename}")

# ---------------- PRETTY PRINT ----------------
def print_rates(data):
    print("\n📊 USD to INR Exchange Rates:")
    for bank, label, rate in data:
        print(f"{bank:<5} | {label:<20} | ₹{rate:.2f}")

# ---------------- MAIN ----------------
if __name__ == "__main__":
    max_retries = 5
    for attempt in range(max_retries):
        if check_internet():
            break
        print(f"🌐 No internet. Retrying in 60s... ({attempt + 1}/{max_retries})")
        time.sleep(60)

    if not check_internet():
        print("❌ Still no internet after retries. Skipping rate collection.")
    else:
        sbi_data = fetch_sbi_all_slabs()
        hdfc_data = fetch_hdfc_usd_cash_buying()
        all_data = sbi_data + hdfc_data

        print_rates(all_data)
        save_rates_to_excel(all_data)
