# gst_downloader_final.py
import os
import time
import shutil
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook

# ---------------- URLs ----------------
LOGIN_URL = "https://services.gst.gov.in/services/login"
MONTHLY_DASHBOARD = "https://return.gst.gov.in/returns/auth/dashboard"
ANNUAL_RETURN_URL = "https://return.gst.gov.in/returns2/auth/annualreturn"

# ---------------- Logging (Excel) ----------------
def init_log_file(folder):
    log_path = os.path.join(folder, "GST_Download_Log.xlsx")
    if not os.path.exists(log_path):
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp", "Financial Year", "Month", "Document", "Status", "File Path"])
        wb.save(log_path)
    return log_path

def log_action(folder, fy, month, document, status, file_path="N/A"):
    log_path = init_log_file(folder)
    wb = load_workbook(log_path)
    ws = wb.active
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([ts, fy, month, document, status, file_path])
    wb.save(log_path)
    print(f"[{ts}] ({fy}) {month} {document} -> {status} | {file_path}")

# ---------------- File helpers ----------------
def get_latest_pdf(download_folder):
    try:
        files = [f for f in os.listdir(download_folder) if f.lower().endswith(".pdf")]
        full_paths = [os.path.join(download_folder, f) for f in files]
        if not full_paths:
            return None
        latest = max(full_paths, key=os.path.getmtime)
        return latest
    except Exception:
        return None

def move_latest_pdf(download_folder, dest_folder):
    latest = get_latest_pdf(download_folder)
    if not latest:
        return None
    os.makedirs(dest_folder, exist_ok=True)
    dest_path = os.path.join(dest_folder, os.path.basename(latest))
    try:
        shutil.move(latest, dest_path)
        return dest_path
    except Exception:
        return None

# ---------------- Selenium helpers ----------------
def wait_for_page_load(driver, timeout=30):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

def safe_click(driver, xpath, wait_time=20):
    """Click element if present and clickable; return True on success."""
    try:
        el = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        el.click()
        return True
    except Exception:
        return False

def element_exists(driver, xpath, wait_time=8):
    try:
        WebDriverWait(driver, wait_time).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False

# ---------------- GUI ----------------
class GSTDownloaderGUI:
    def __init__(self, master):
        self.master = master
        master.title("GST PDF Downloader")

        # Username
        tk.Label(master, text="Username").grid(row=0, column=0, sticky="e")
        self.username_entry = tk.Entry(master, width=40)
        self.username_entry.grid(row=0, column=1, padx=6, pady=4)

        # Password
        tk.Label(master, text="Password").grid(row=1, column=0, sticky="e")
        self.password_entry = tk.Entry(master, width=40, show="*")
        self.password_entry.grid(row=1, column=1, padx=6, pady=4)

        # Destination folder
        tk.Label(master, text="Destination Folder").grid(row=2, column=0, sticky="e")
        self.dest_var = tk.StringVar()
        tk.Entry(master, textvariable=self.dest_var, width=36).grid(row=2, column=1, sticky="w")
        tk.Button(master, text="Browse", command=self.browse_dest).grid(row=2, column=2, padx=6)

        # ChromeDriver path
        tk.Label(master, text="ChromeDriver Path").grid(row=3, column=0, sticky="e")
        self.driver_var = tk.StringVar()
        tk.Entry(master, textvariable=self.driver_var, width=36).grid(row=3, column=1, sticky="w")
        tk.Button(master, text="Browse", command=self.browse_driver).grid(row=3, column=2, padx=6)

        # FY checkboxes
        tk.Label(master, text="Select Financial Years").grid(row=4, column=0, sticky="ne")
        self.fy_vars = {}
        years = ["FY 2017-18", "FY 2018-19", "FY 2019-20", "FY 2020-21",
                 "FY 2021-22", "FY 2022-23", "FY 2023-24", "FY 2024-25"]
        for i, y in enumerate(years):
            var = tk.BooleanVar()
            cb = tk.Checkbutton(master, text=y, variable=var)
            cb.grid(row=4 + (i // 4), column=1 + (i % 4), sticky="w")
            self.fy_vars[y] = var

        # Submit
        tk.Button(master, text="Submit", command=self.submit, width=18, bg="#1976D2", fg="white").grid(row=7, column=1, pady=12)

    def browse_dest(self):
        d = filedialog.askdirectory()
        if d:
            self.dest_var.set(d)

    def browse_driver(self):
        p = filedialog.askopenfilename(filetypes=[("ChromeDriver executable", "*chromedriver*;*.exe")])
        if p:
            self.driver_var.set(p)

    def submit(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        dest = self.dest_var.get().strip()
        driver_path = self.driver_var.get().strip()
        selected_fys = [fy for fy, v in self.fy_vars.items() if v.get()]

        if not username or not password:
            messagebox.showerror("Error", "Username and Password required.")
            return
        if not dest:
            messagebox.showerror("Error", "Please select destination folder.")
            return
        if not driver_path:
            messagebox.showerror("Error", "Please select ChromeDriver path.")
            return
        if not selected_fys:
            messagebox.showerror("Error", "Please select at least one Financial Year.")
            return

        self.master.destroy()
        run_automation(username, password, dest, driver_path, selected_fys)

# ---------------- Main automation ----------------
def run_automation(username, password, base_folder, chrome_driver_path, selected_fys):
    # Setup chrome options to auto-download PDFs
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": base_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_settings.popups": 0
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--safebrowsing-disable-download-protection")

    # Use provided chromedriver path
    service = Service(chrome_driver_path)

    driver = None
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()

        # go to login
        driver.get(LOGIN_URL)
        wait_for_page_load(driver)
        # Fill username and password
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username"))).clear()
            driver.find_element(By.ID, "username").send_keys(username)
        except Exception:
            # fallback: try by name attribute
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "user_name")))
                driver.find_element(By.NAME, "user_name").send_keys(username)
            except Exception:
                log_action(base_folder, "N/A", "N/A", "Login - username field not found", "Fail")
                return

        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "user_pass"))).clear()
            driver.find_element(By.ID, "user_pass").send_keys(password)
        except Exception:
            log_action(base_folder, "N/A", "N/A", "Login - password field not found", "Fail")
            return

        # Wait for manual captcha entry and user to press Login in browser.
        # We wait for the URL to change to something containing 'dashboard' (max 5 minutes).
        print("Please enter CAPTCHA manually in the opened browser window. Waiting for you to login...")
        try:
            WebDriverWait(driver, 300).until(EC.url_contains("dashboard"))
            wait_for_page_load(driver)
            log_action(base_folder, "N/A", "N/A", "Login", "Success")
        except Exception as e:
            log_action(base_folder, "N/A", "N/A", f"Login did not reach dashboard: {e}", "Fail")
            return

        # For each selected FY do entire sequence (monthly GSTR-1 -> GSTR-3B -> annual -> 9C)
        for fy_label in selected_fys:
            # Use FY label text for selects (site may expect "2018-19" etc.)
            fy_text = fy_label.replace("FY ", "").strip() if fy_label.startswith("FY") else fy_label

            # --- MONTHLY: GSTR-1 (iterate months if month dropdown exists) ---
            try:
                driver.get(MONTHLY_DASHBOARD)
                wait_for_page_load(driver)
                # select financial year in monthly dashboard (select[name='fin'])
                try:
                    fin_sel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "fin")))
                    Select(fin_sel).select_by_visible_text(fy_text)
                except Exception:
                    # log but continue
                    log_action(base_folder, fy_label, "N/A", "Select FY on monthly dashboard", "Fail")

                # Try to collect months from select[name='mon'] if present
                months = []
                try:
                    msel = driver.find_element(By.NAME, "mon")
                    mopts = Select(msel).options
                    months = [opt.text.strip() for opt in mopts if opt.text.strip()]
                    if not months:
                        months = [None]
                except Exception:
                    months = [None]

                for month in months:
                    month_label = month if month else "N/A"
                    # select month if applicable
                    if month:
                        try:
                            Select(driver.find_element(By.NAME, "mon")).select_by_visible_text(month)
                        except Exception:
                            pass

                    # Click Search (button.srchbtn or button[type=submit] with text Search)
                    search_clicked = False
                    search_xpaths = [
                        "//button[contains(@class,'srchbtn') and (contains(.,'Search') or contains(@data-ng-bind,'trans.LBL_SCH'))]",
                        "//button[@type='submit' and contains(normalize-space(.),'Search')]",
                        "//button[contains(@class,'srchbtn')]"
                    ]
                    for sx in search_xpaths:
                        if safe_click(driver, sx):
                            search_clicked = True
                            break
                    if not search_clicked:
                        log_action(base_folder, fy_label, month_label, "Search (monthly)", "Fail")
                        # still continue to try GSTR-1 flow once
                    else:
                        wait_for_page_load(driver)
                        log_action(base_folder, fy_label, month_label, "Search (monthly)", "Success")

                    # Click VIEW inside GSTR-1 element box
                    # Use the provided element structure: data-ng-if check for GSTR1 and a button with text VIEW
                    gstr1_view_xpath = "//div[@data-ng-if=\"x.return_ty=='GSTR1'\" or contains(.,'GSTR-1')]//button[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'VIEW')]"
                    if safe_click(driver, gstr1_view_xpath):
                        wait_for_page_load(driver)
                        log_action(base_folder, fy_label, month_label, "GSTR-1: Click View", "Success")
                        # Click View Summary button inside GSTR-1 page
                        gstr1_view_summary_xpath = "//button[.//span[contains(normalize-space(.),'VIEW SUMMARY') or contains(normalize-space(.),'PROCEED TO FILE/SUMMARY')]]"
                        if safe_click(driver, gstr1_view_summary_xpath):
                            wait_for_page_load(driver)
                            log_action(base_folder, fy_label, month_label, "GSTR-1: View Summary", "Success")
                        else:
                            log_action(base_folder, fy_label, month_label, "GSTR-1: View Summary", "Fail")

                        # Click Download (PDF) button in GSTR-1 summary (data-ng-click='genratepdfNew()' or visible text)
                        gstr1_download_xpath = "//button[contains(@data-ng-click,'genratepdfNew') or contains(normalize-space(.),'DOWNLOAD (PDF)') or contains(normalize-space(.),'DOWNLOAD SUMMARY (PDF)')]"
                        if safe_click(driver, gstr1_download_xpath):
                            # move latest pdf into <base>/<FY>/GSTR-1/
                            dest = os.path.join(base_folder, fy_label, "GSTR-1")
                            moved = move_latest_pdf(base_folder, dest)
                            if moved:
                                log_action(base_folder, fy_label, month_label, "Download GSTR-1 (PDF)", "Success", moved)
                            else:
                                log_action(base_folder, fy_label, month_label, "Download GSTR-1 (PDF)", "Success", "File not found")
                        else:
                            log_action(base_folder, fy_label, month_label, "Download GSTR-1 (PDF)", "Fail")
                    else:
                        log_action(base_folder, fy_label, month_label, "GSTR-1: Click View", "Fail")

                # End month loop
            except Exception as e:
                log_action(base_folder, fy_label, "N/A", f"GSTR-1 overall error: {e}", "Fail")

            # --- MONTHLY: GSTR-3B (direct download after Search) ---
            try:
                driver.get(MONTHLY_DASHBOARD)
                wait_for_page_load(driver)
                try:
                    fin_sel = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "fin")))
                    Select(fin_sel).select_by_visible_text(fy_text)
                except Exception:
                    log_action(base_folder, fy_label, "N/A", "Select FY on monthly dashboard (GSTR-3B)", "Fail")

                # months as before
                months = []
                try:
                    msel = driver.find_element(By.NAME, "mon")
                    months = [opt.text.strip() for opt in Select(msel).options if opt.text.strip()]
                    if not months:
                        months = [None]
                except Exception:
                    months = [None]

                for month in months:
                    month_label = month if month else "N/A"
                    if month:
                        try:
                            Select(driver.find_element(By.NAME, "mon")).select_by_visible_text(month)
                        except Exception:
                            pass

                    # click search
                    search_clicked = False
                    for sx in ["//button[contains(@class,'srchbtn')]", "//button[contains(normalize-space(.),'Search')]"]:
                        if safe_click(driver, sx):
                            search_clicked = True
                            break
                    if not search_clicked:
                        log_action(base_folder, fy_label, month_label, "Search (GSTR-3B)", "Fail")
                        continue
                    wait_for_page_load(driver)
                    log_action(base_folder, fy_label, month_label, "Search (GSTR-3B)", "Success")

                    # click download button for GSTR-3B (data-ng-click='downloadGSTR3Bpdf()')
                    gstr3b_xpath = "//button[@data-ng-click='downloadGSTR3Bpdf()' or contains(normalize-space(.),'Download')]"
                    if safe_click(driver, gstr3b_xpath):
                        dest = os.path.join(base_folder, fy_label, "GSTR-3B")
                        moved = move_latest_pdf(base_folder, dest)
                        if moved:
                            log_action(base_folder, fy_label, month_label, "Download GSTR-3B (PDF)", "Success", moved)
                        else:
                            log_action(base_folder, fy_label, month_label, "Download GSTR-3B (PDF)", "Success", "File not found")
                    else:
                        log_action(base_folder, fy_label, month_label, "Download GSTR-3B (PDF)", "Fail")
            except Exception as e:
                log_action(base_folder, fy_label, "N/A", f"GSTR-3B overall error: {e}", "Fail")

            # --- ANNUAL: GSTR-9, GSTR-1 Annual, GSTR-3B Annual ---
            try:
                driver.get(ANNUAL_RETURN_URL)
                wait_for_page_load(driver)
                # select financial year (select[name='finyr'])
                try:
                    finyr_el = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "finyr")))
                    Select(finyr_el).select_by_visible_text(fy_text)
                except Exception:
                    log_action(base_folder, fy_label, "N/A", "Select FY on annual page", "Fail")

                # click Search
                if safe_click(driver, "//button[contains(@class,'srchbtn') or contains(normalize-space(.),'Search')]"):
                    wait_for_page_load(driver)
                    log_action(base_folder, fy_label, "N/A", "Annual Search", "Success")
                else:
                    log_action(base_folder, fy_label, "N/A", "Annual Search", "Fail")

                # Click View GSTR-9 (button text 'VIEW GSTR-9' or page_rtp)
                gstr9_view_candidates = [
                    "//button[contains(text(),'VIEW GSTR-9')]",
                    "//button[@data-ng-click='page_rtp(x.return_ty,x.due_dt,x.status)' and contains(.,'GSTR-9')]",
                    "//button[contains(@data-ng-click,'page_rtp') and contains(.,'GSTR-9')]"
                ]
                viewed = False
                for xp in gstr9_view_candidates:
                    if safe_click(driver, xp):
                        viewed = True
                        wait_for_page_load(driver)
                        break
                if not viewed:
                    log_action(base_folder, fy_label, "N/A", "View GSTR-9 not found", "Fail")
                else:
                    # Inside GSTR-9 view: attempt downloads in required order

                    # Download GSTR-9 details (PDF)
                    if safe_click(driver, "//button[contains(@data-ng-click,'getPdfData_gstr9') or contains(normalize-space(.),'Download GSTR-9')]"):
                        dest = os.path.join(base_folder, fy_label, "GSTR-9")
                        moved = move_latest_pdf(base_folder, dest)
                        if moved:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-9 (PDF)", "Success", moved)
                        else:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-9 (PDF)", "Success", "File not found")
                    else:
                        log_action(base_folder, fy_label, "N/A", "Download GSTR-9 (PDF)", "Fail")

                    # Download GSTR-1 Annual / IFF Summary
                    if safe_click(driver, "//button[contains(@data-ng-click,'getPdfData_gstr1') or contains(normalize-space(.),'GSTR-1/IFF SUMMARY')]"):
                        dest = os.path.join(base_folder, fy_label, "GSTR-1 Annual")
                        moved = move_latest_pdf(base_folder, dest)
                        if moved:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-1 Annual (PDF)", "Success", moved)
                        else:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-1 Annual (PDF)", "Success", "File not found")
                    else:
                        log_action(base_folder, fy_label, "N/A", "Download GSTR-1 Annual (PDF)", "Fail")

                    # Download GSTR-3B Annual
                    if safe_click(driver, "//button[contains(@data-ng-click,'getPdfData_gstr3B') or contains(normalize-space(.),'GSTR-3B SUMMARY')]"):
                        dest = os.path.join(base_folder, fy_label, "GSTR-3B Annual")
                        moved = move_latest_pdf(base_folder, dest)
                        if moved:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-3B Annual (PDF)", "Success", moved)
                        else:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-3B Annual (PDF)", "Success", "File not found")
                    else:
                        log_action(base_folder, fy_label, "N/A", "Download GSTR-3B Annual (PDF)", "Fail")

            except Exception as e:
                log_action(base_folder, fy_label, "N/A", f"Annual return overall error: {e}", "Fail")

            # --- GSTR-9C ---
            try:
                # Click the Download GSTR-9C button which triggers offlinepath and navigates to new page
                gstr9c_btn_xpath = "//button[contains(@data-ng-click,'offlinepath') or contains(normalize-space(.),'DOWNLOAD GSTR-9C')]"
                if safe_click(driver, gstr9c_btn_xpath):
                    # wait for navigation (URL change) and page load
                    old_url = driver.current_url
                    try:
                        WebDriverWait(driver, 20).until(lambda d: d.current_url != old_url)
                    except Exception:
                        pass
                    wait_for_page_load(driver)
                    # Now click 'Download Filed GSTR-9C(PDF)' button (data-ng-click='generate9cpdf()')
                    if safe_click(driver, "//button[@data-ng-click='generate9cpdf()' or contains(normalize-space(.),'Download filed GSTR-9C')]"):
                        dest = os.path.join(base_folder, fy_label, "GSTR-9C")
                        moved = move_latest_pdf(base_folder, dest)
                        if moved:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-9C (PDF)", "Success", moved)
                        else:
                            log_action(base_folder, fy_label, "N/A", "Download GSTR-9C (PDF)", "Success", "File not found")
                    else:
                        log_action(base_folder, fy_label, "N/A", "Download GSTR-9C (PDF)", "Fail")
                else:
                    log_action(base_folder, fy_label, "N/A", "GSTR-9C offline button not found", "Fail")
            except Exception as e:
                log_action(base_folder, fy_label, "N/A", f"GSTR-9C overall error: {e}", "Fail")

        # end per-FY loop

    except Exception as ex_all:
        log_action(base_folder, "N/A", "N/A", f"Unhandled exception: {ex_all}", "Fail")
        print("Unhandled exception:", traceback.format_exc())
    finally:
        # Ensure browser cleaned up
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

        # After run, open summary window showing failures
        show_results(base_folder)

# ---------------- Results window ----------------
def show_results(folder):
    log_path = init_log_file(folder)
    wb = load_workbook(log_path)
    ws = wb.active
    failures = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        ts, fy, month, doc, status, fpath = row
        if status and "Fail" in str(status):
            failures.append((fy or "N/A", month or "N/A", doc or "N/A"))

    root = tk.Tk()
    root.title("Download Results")
    root.geometry("600x400")

    if not failures:
        tk.Label(root, text="✅ All downloads completed successfully!", fg="green", font=("Arial", 14)).pack(pady=20)
    else:
        tk.Label(root, text="❌ Some downloads failed:", fg="red", font=("Arial", 14)).pack(pady=8)
        cols = ("Financial Year", "Month", "Document")
        tree = ttk.Treeview(root, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=180 if c != "Document" else 220)
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for r in failures:
            tree.insert("", "end", values=r)

    tk.Button(root, text="Close", command=root.destroy).pack(pady=8)
    root.mainloop()

# ---------------- Main ----------------
if __name__ == "__main__":
    app_root = tk.Tk()
    gui = GSTDownloaderGUI(app_root)
    app_root.mainloop()
