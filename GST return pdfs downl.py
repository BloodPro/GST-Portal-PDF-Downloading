import os
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime

# ================== GUI ================== #
class GSTDownloaderGUI:
    def __init__(self, master):
        self.master = master
        master.title("GST PDF Downloader")

        # Username
        tk.Label(master, text="Username").grid(row=0, column=0, sticky="w")
        self.username_entry = tk.Entry(master, width=40)
        self.username_entry.grid(row=0, column=1, pady=2)

        # Password
        tk.Label(master, text="Password").grid(row=1, column=0, sticky="w")
        self.password_entry = tk.Entry(master, width=40, show="*")
        self.password_entry.grid(row=1, column=1, pady=2)

        # Destination Folder
        tk.Label(master, text="Destination Folder").grid(row=2, column=0, sticky="w")
        self.dest_folder_entry = tk.Entry(master, width=40)
        self.dest_folder_entry.grid(row=2, column=1, pady=2)
        tk.Button(master, text="Browse", command=self.browse_dest_folder).grid(row=2, column=2, padx=5)

        # Chrome WebDriver Path
        tk.Label(master, text="Chrome WebDriver Path").grid(row=3, column=0, sticky="w")
        self.driver_path_entry = tk.Entry(master, width=40)
        self.driver_path_entry.grid(row=3, column=1, pady=2)
        tk.Button(master, text="Browse", command=self.browse_driver_path).grid(row=3, column=2, padx=5)

        # FY Checkboxes
        self.fy_vars = {}
        fy_years = [
            "2017-18", "2018-19", "2019-20", "2020-21",
            "2021-22", "2022-23", "2023-24", "2024-25"
        ]
        tk.Label(master, text="Select FYs:").grid(row=4, column=0, sticky="w")
        for i, fy in enumerate(fy_years):
            var = tk.BooleanVar()
            cb = tk.Checkbutton(master, text=fy, variable=var)
            cb.grid(row=5 + i // 4, column=i % 4, sticky="w")
            self.fy_vars[fy] = var

        # Submit Button
        tk.Button(master, text="Submit", command=self.submit).grid(row=8, column=1, pady=10)

    def browse_dest_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dest_folder_entry.delete(0, tk.END)
            self.dest_folder_entry.insert(0, folder)

    def browse_driver_path(self):
        path = filedialog.askopenfilename(filetypes=[("ChromeDriver", "*.exe")])
        if path:
            self.driver_path_entry.delete(0, tk.END)
            self.driver_path_entry.insert(0, path)

    def submit(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        dest_folder = self.dest_folder_entry.get()
        driver_path = self.driver_path_entry.get()
        selected_fys = [fy for fy, var in self.fy_vars.items() if var.get()]

        if not username or not password or not dest_folder or not driver_path or not selected_fys:
            messagebox.showerror("Error", "Please fill all fields and select at least one FY.")
            return

        self.master.destroy()
        run_gst_download(username, password, dest_folder, driver_path, selected_fys)

# ================== Selenium Automation ================== #
def run_gst_download(username, password, dest_folder, driver_path, selected_fys):
    log_data = []
    failed_docs = []
    driver = None

    try:
        chrome_options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": dest_folder,
            "plugins.always_open_pdf_externally": True,
            "download.prompt_for_download": False
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--start-maximized")

        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.get("https://services.gst.gov.in/services/login")

        # Login
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "user_pass").send_keys(password)
        driver.find_element(By.ID, "loginBtn").click()

        # Loop through selected FYs
        for fy in selected_fys:
            try:
                # === GSTR-1 Annual Download ===
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='getPdfData_gstr1()']"))).click()
                log_data.append({"FY": fy, "Month": "", "Document": "GSTR-1 Annual", "Path": dest_folder})

                # === GSTR-3B Annual Download ===
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='getPdfData_gstr3B()']"))).click()
                log_data.append({"FY": fy, "Month": "", "Document": "GSTR-3B Annual", "Path": dest_folder})

                # === GSTR-9 ===
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='page_rtp(x.return_ty,x.due_dt,x.status)']"))).click()
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='getPdfData_gstr9()']"))).click()
                log_data.append({"FY": fy, "Month": "", "Document": "GSTR-9", "Path": dest_folder})

                # === GSTR-9C ===
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='offlinepath(x.return_ty,x.status)']"))).click()
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[@data-ng-click='generate9cpdf()']"))).click()
                log_data.append({"FY": fy, "Month": "", "Document": "GSTR-9C", "Path": dest_folder})

            except Exception as e:
                failed_docs.append({"FY": fy, "Month": "", "Document": "One or more"})
                print(f"Failed for {fy}: {e}")

    except Exception:
        print("Error:", traceback.format_exc())
    finally:
        if driver:
            driver.quit()

        # Save log to Excel
        log_file_path = os.path.join(dest_folder, f"GST_Download_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        pd.DataFrame(log_data).to_excel(log_file_path, index=False)

        # Show summary popup
        root = tk.Tk()
        root.withdraw()
        if failed_docs:
            fail_text = "\n".join([f"{f['FY']} - {f['Month']} - {f['Document']}" for f in failed_docs])
            messagebox.showerror("Download Completed with Failures", f"Some documents failed:\n\n{fail_text}")
        else:
            messagebox.showinfo("Success", f"All documents downloaded successfully.\nLog saved at:\n{log_file_path}")

# ================== Main ================== #
if __name__ == "__main__":
    root = tk.Tk()
    app = GSTDownloaderGUI(root)
    root.mainloop()
