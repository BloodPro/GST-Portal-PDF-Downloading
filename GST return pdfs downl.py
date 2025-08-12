import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook
from datetime import datetime

# ---------------- Logging Functions ----------------
def init_log_file(folder):
    log_path = os.path.join(folder, "GST_Download_Log.xlsx")
    if not os.path.exists(log_path):
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp", "Financial Year", "Action", "Status"])
        wb.save(log_path)
    return log_path

def log_action(folder, fy, action, status):
    log_path = init_log_file(folder)
    wb = load_workbook(log_path)
    ws = wb.active
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([timestamp, fy, action, status])
    wb.save(log_path)
    print(f"[{timestamp}] ({fy}) {action} → {status}")

# ---------------- Helper Functions ----------------
def safe_click(driver, folder, fy, description, xpath, wait_time=20):
    """Wait for element to be clickable and click it."""
    try:
        WebDriverWait(driver, wait_time).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        ).click()
        time.sleep(2)
        log_action(folder, fy, description, "Success")
        return True
    except Exception as e:
        log_action(folder, fy, description, f"Fail ({e})")
        return False

def wait_for_page_load(driver, folder, fy, description, timeout=20):
    """Wait until document.readyState is complete."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        log_action(folder, fy, description, "Success")
    except Exception as e:
        log_action(folder, fy, description, f"Fail ({e})")

def wait_for_url_change(driver, old_url, folder, fy, description, timeout=20):
    """Wait for URL to change after an action."""
    try:
        WebDriverWait(driver, timeout).until(lambda d: d.current_url != old_url)
        log_action(folder, fy, description, "Success")
        return True
    except Exception as e:
        log_action(folder, fy, description, f"Fail ({e})")
        return False

# ---------------- GUI ----------------
class GSTDownloaderGUI:
    def __init__(self, master):
        self.master = master
        master.title("GST PDF Downloader")

        # Username
        tk.Label(master, text="Username").grid(row=0, column=0, sticky="e")
        self.username_entry = tk.Entry(master, width=30)
        self.username_entry.grid(row=0, column=1)

        # Password
        tk.Label(master, text="Password").grid(row=1, column=0, sticky="e")
        self.password_entry = tk.Entry(master, width=30, show="*")
        self.password_entry.grid(row=1, column=1)

        # Destination folder
        tk.Label(master, text="Destination Folder").grid(row=2, column=0, sticky="e")
        self.folder_path = tk.StringVar()
        tk.Entry(master, textvariable=self.folder_path, width=30).grid(row=2, column=1)
        tk.Button(master, text="Browse", command=self.browse_folder).grid(row=2, column=2)

        # Financial year checkboxes
        tk.Label(master, text="Select Financial Years").grid(row=3, column=0, sticky="ne")
        self.years_vars = {}
        years = ["2017-18", "2018-19", "2019-20", "2020-21", "2021-22", "2022-23", "2023-24", "2024-25"]
        for i, year in enumerate(years):
            var = tk.BooleanVar()
            chk = tk.Checkbutton(master, text=year, variable=var)
            chk.grid(row=3 + i // 2, column=1 + (i % 2))
            self.years_vars[year] = var

        # Submit
        tk.Button(master, text="Submit", command=self.submit).grid(row=8, column=1, pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)

    def submit(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        folder = self.folder_path.get().strip()
        years = [y for y, var in self.years_vars.items() if var.get()]

        if not username or not password or not folder or not years:
            messagebox.showerror("Error", "Please fill in all fields and select at least one financial year.")
            return

        self.master.destroy()
        self.result = {
            "username": username,
            "password": password,
            "folder": folder,
            "years": years
        }

# ---------------- Selenium Automation ----------------
def gst_download(username, password, folder, years):
    chrome_options = Options()
    prefs = {"download.default_directory": folder, "plugins.always_open_pdf_externally": True}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)

    # Login
    driver.get("https://services.gst.gov.in/services/login")
    wait_for_page_load(driver, folder, "Login", "Login Page Load")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
    driver.find_element(By.ID, "user_pass").send_keys(password)

    print("Please enter captcha manually and log in...")
    WebDriverWait(driver, 120).until(EC.url_contains("dashboard"))
    wait_for_page_load(driver, folder, "Login", "Dashboard Loaded")

    for fy in years:
        print(f"Processing year: {fy}")

        # ---------------- MONTHLY GSTR-1 ----------------
        driver.get("https://return.gst.gov.in/returns/auth/dashboard")
        wait_for_page_load(driver, folder, fy, "GSTR-1 Dashboard Loaded")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "financialYear"))).send_keys(fy)
        safe_click(driver, folder, fy, "Search GSTR-1", "//button[@id='searchButton']")
        wait_for_page_load(driver, folder, fy, "GSTR-1 Search Results Loaded")
        safe_click(driver, folder, fy, "View GSTR-1", "//button[contains(text(),'VIEW')]")
        wait_for_page_load(driver, folder, fy, "GSTR-1 View Page Loaded")
        safe_click(driver, folder, fy, "View Summary GSTR-1", "//span[contains(text(),'VIEW SUMMARY')]")
        wait_for_page_load(driver, folder, fy, "GSTR-1 Summary Loaded")
        safe_click(driver, folder, fy, "Download GSTR-1 PDF", "//span[contains(text(),'DOWNLOAD (PDF)')]")

        # ---------------- MONTHLY GSTR-3B ----------------
        driver.get("https://return.gst.gov.in/returns/auth/dashboard")
        wait_for_page_load(driver, folder, fy, "GSTR-3B Dashboard Loaded")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "financialYear"))).send_keys(fy)
        safe_click(driver, folder, fy, "Search GSTR-3B", "//button[@id='searchButton']")
        wait_for_page_load(driver, folder, fy, "GSTR-3B Search Results Loaded")
        safe_click(driver, folder, fy, "Download GSTR-3B PDF", "//button[@data-ng-click='downloadGSTR3Bpdf()']")

        # ---------------- ANNUAL RETURN ----------------
        driver.get("https://services.gst.gov.in/services/annualreturn")
        wait_for_page_load(driver, folder, fy, "Annual Return Page Loaded")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "finyr")))
        Select(driver.find_element(By.NAME, "finyr")).select_by_visible_text(fy)
        safe_click(driver, folder, fy, "Search Annual Return", "//button[@type='submit' and contains(@class,'srchbtn')]")
        wait_for_page_load(driver, folder, fy, "Annual Return Search Results Loaded")

        # View GSTR-9 & Download PDFs
        if safe_click(driver, folder, fy, "View GSTR-9", "//button[@data-ng-click='page_rtp(x.return_ty,x.due_dt,x.status)']"):
            wait_for_page_load(driver, folder, fy, "GSTR-9 View Page Loaded")
            safe_click(driver, folder, fy, "Download Annual GSTR-1 PDF", "//button[@data-ng-click='getPdfData_gstr1()']")
            safe_click(driver, folder, fy, "Download Annual GSTR-3B PDF", "//button[@data-ng-click='getPdfData_gstr3B()']")
            safe_click(driver, folder, fy, "Download GSTR-9 PDF", "//button[@data-ng-click='getPdfData_gstr9()']")

        # ---------------- GSTR-9C ----------------
        if safe_click(driver, folder, fy, "Click Download GSTR-9C Button", "//button[@data-ng-click='offlinepath(x.return_ty,x.status)']"):
            old_url = driver.current_url
            wait_for_url_change(driver, old_url, folder, fy, "Navigated to GSTR-9C Page")
            wait_for_page_load(driver, folder, fy, "GSTR-9C Page Loaded")
            safe_click(driver, folder, fy, "Download GSTR-9C PDF", "//button[@data-ng-click='generate9cpdf()']")

    driver.quit()

# ---------------- Main ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = GSTDownloaderGUI(root)
    root.mainloop()

    if hasattr(app, "result"):
        data = app.result
        gst_download(data["username"], data["password"], data["folder"], data["years"])
