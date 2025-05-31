import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import shutil
import time
import pyautogui
from pathlib import Path
from datetime import datetime


class FiltererApp:
    def __init__(self, root):
        self.root = root
        self.root.title("0 Filterer (Local + Pandas)")

        # ------------------------ In-Memory Data ------------------------
        self.df_inventory = pd.DataFrame()  # Will hold your full Excel in memory
        self.filtered_zeros = set()         # Articles with inventory <= 0 (not in DNO)
        self.filtered_lows = set()          # Articles with 0 < inventory <= threshold
        self.zero_count = 0                # For UI display
        self.new_found_dnos = 0
        self.inputted = False              # For optional logging usage

        # Departments + lights
        self.departments = {
            "Grocery": ["Grocery"],
            "Meat": ["Meat", "Deli"],
            "Bakery": ["Bakery Commercial", "Bakery Instore"],
            "Dairy/Frozen": ["Bulk"],
            "Seafood": ["Seafood"],
            "HMR": ["HMR"],
            "Produce": ["Produce"],
            "Home": ["Home", "Entertainment"]
        }

        # Banned categories (not to be reported)
        self.BANNED_CATS = [
            "Nuts/ Dried Fruit", "Fresh-", "Field Veg", "Root Veg", "Salad Veg",
            "Cooking Veg", "Peppers", "Tomatoes",  # produce
            "Lamb", "Sausage", "Hams",            # meat
            "Books-", "Magazines", "Newspapers"   # entertainment
        ]

        self.lights_bool = {dep: False for dep in self.departments}
        self.LOW_THRESHOLD = 2  # Hyperparameter for "low" inventory

        # ------------------------ UI LAYOUT ------------------------

        # 1) DEPARTMENT LIGHTS
        self.dept_frame = tk.Frame(root)
        self.dept_frame.grid_columnconfigure((0, 2), weight=1)
        self.dept_frame.pack(pady=20, padx=20)

        self.buttons = {}
        self.lights = {}
        for idx, key in enumerate(self.departments):
            lbl = tk.Label(self.dept_frame, text=key)
            lbl.grid(row=idx // 2, column=2 * (idx % 2), padx=10, pady=5)
            self.buttons[key] = lbl

            light = tk.Canvas(self.dept_frame, width=20, height=20)
            light.create_oval(2, 2, 18, 18, fill="red", tags="light")
            light.grid(row=idx // 2, column=2 * (idx % 2) + 1)
            self.lights[key] = light

        # 2) MAIN FUNCTION FRAME
        self.fxn_frame = tk.Frame(root)
        self.fxn_frame.pack(pady=20, padx=20)

        # We'll create sub-frames to group related controls
        # (A) DNO File Management
        self.dno_file_frame = tk.LabelFrame(self.fxn_frame, text="DNO File Management", padx=10, pady=10)
        self.dno_file_frame.grid(row=0, column=0, sticky="n", padx=10)

        self.import_dno_btn = tk.Button(self.dno_file_frame, text="Import DNO File", command=self.import_dno)
        self.import_dno_btn.pack(padx=5, pady=(0, 5))

        self.export_dno_btn = tk.Button(self.dno_file_frame, text="Export DNO File", command=self.export_dno)
        self.export_dno_btn.pack(padx=5, pady=5)

        # (B) Single-article DNO Management
        self.dno_single_frame = tk.LabelFrame(self.fxn_frame, text="Single-Article DNO", padx=10, pady=10)
        self.dno_single_frame.grid(row=0, column=1, sticky="n", padx=10)

        tk.Label(self.dno_single_frame, text="Article #:").pack()
        self.entry = tk.Entry(self.dno_single_frame)
        self.entry.pack(pady=(0, 10))

        self.add_ONE_btn = tk.Button(self.dno_single_frame, text="Add to DNO", command=self.add_new_DNO)
        self.add_ONE_btn.pack(padx=5, pady=2, fill=tk.X)

        self.remove_ONE_btn = tk.Button(self.dno_single_frame, text="Remove from DNO", command=self.remove_from_DNO)
        self.remove_ONE_btn.pack(padx=5, pady=2, fill=tk.X)

        # (C) Inventory & Filtering
        self.inv_frame = tk.LabelFrame(self.fxn_frame, text="Inventory & Filters", padx=10, pady=10)
        self.inv_frame.grid(row=0, column=2, sticky="n", padx=10)

        self.upload_btn = tk.Button(self.inv_frame, text="Upload Excel", command=self.upload_excel)
        self.upload_btn.pack(padx=5, pady=5, fill=tk.X)

        self.find_zeros_btn = tk.Button(self.inv_frame, text="Find Zeros", command=self.find_zeros)
        self.find_zeros_btn.pack(padx=5, pady=5, fill=tk.X)

        self.find_lows_btn = tk.Button(self.inv_frame, text="Find Lows", command=self.find_lows)
        self.find_lows_btn.pack(padx=5, pady=5, fill=tk.X)

        # We'll dynamically add "Send Zeros/Lows to SAP" buttons below these
        self.zero_button = None
        self.low_button = None

        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    # ------------------------ DNO DB Methods ------------------------

    def fetch_dno_articles(self):
        """
        Returns a set of article IDs from the local dno.db.
        """
        if not os.path.exists("dno.db"):
            return set()
        conn = sqlite3.connect("dno.db")
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS dno (article TEXT UNIQUE)")
        cur.execute("SELECT article FROM dno")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return [int(r[0]) for r in rows]  # storing articles as strings

    def import_dno(self):
        """
        Import an Excel or DB file to replace/merge into local dno.db.
        """
        file_path = filedialog.askopenfilename(
            title="Select the dno file",
            filetypes=[("Excel Files", "*.xlsx"), ("SQLite DB", "*.db")]
        )
        if not file_path:
            self.show_alert("Error", "No file selected.")
            return

        if file_path.endswith(".db"):
            # Replace local dno.db
            shutil.copyfile(file_path, "dno.db")
            self.show_alert("Success", "dno.db replaced.")
        else:
            # Excel
            try:
                xls = pd.read_excel(file_path, sheet_name=None)
                all_values = []
                for _, df in xls.items():
                    vals = pd.Series(df.iloc[:, :10].values.ravel()).dropna()
                    all_values.extend(vals)

                conn = sqlite3.connect("dno.db")
                cur = conn.cursor()
                cur.execute("CREATE TABLE IF NOT EXISTS dno (article TEXT UNIQUE)")

                existing_count = cur.execute("SELECT COUNT(*) FROM dno").fetchone()[0]

                for item in all_values:
                    if isinstance(item, (int, float)):
                        item = str(int(item))
                    cur.execute("INSERT OR IGNORE INTO dno (article) VALUES (?)", (str(item),))

                conn.commit()
                new_count = cur.execute("SELECT COUNT(*) FROM dno").fetchone()[0]
                conn.close()

                diff = new_count - existing_count
                if diff > 0:
                    self.show_alert("", f"Imported {diff} new DNO articles.")
                    self.new_found_dnos += diff
                else:
                    self.show_alert("", "No new DNO articles or all duplicates.")
            except Exception as e:
                self.show_alert("Error", f"An error occurred importing DNO:\n{e}")

    def export_dno(self):
        """
        Exports the current dno.db to the user's Downloads folder.
        """
        downloads_path = Path.home() / "Downloads"
        src = "dno.db"
        dst = downloads_path / "dno.db"

        if not os.path.exists(src):
            self.show_alert("Error", "No local dno.db found to export.")
            return

        try:
            shutil.copyfile(src, dst)
            self.show_alert("Success", f"dno.db exported to:\n{dst}")
        except Exception as e:
            self.show_alert("Error", f"Unable to export dno.db: {e}")

    def add_new_DNO(self):
        article = self.entry.get().strip()
        if not article:
            return
        confirm = messagebox.askokcancel(
            "Confirm Action",
            f"Insert {article} into DNO?\nDouble-check article number."
        )
        if confirm:
            conn = sqlite3.connect("dno.db")
            cur = conn.cursor()
            try:
                cur.execute("CREATE TABLE IF NOT EXISTS dno (article TEXT UNIQUE)")
                cur.execute("INSERT OR IGNORE INTO dno (article) VALUES (?)", (article,))
                conn.commit()
                self.show_alert("Inserted", f"Article {article} added to DNO.")
                self.new_found_dnos += 1
            except sqlite3.Error as e:
                self.show_alert("Error", str(e))
            finally:
                conn.close()

    def remove_from_DNO(self):
        article = self.entry.get().strip()
        if not article:
            return
        confirm = messagebox.askokcancel(
            "Confirm Action",
            f"Remove {article} from DNO?\nDouble-check article number."
        )
        if confirm:
            conn = sqlite3.connect("dno.db")
            cur = conn.cursor()
            try:
                cur.execute("DELETE FROM dno WHERE article = ?", (article,))
                conn.commit()
                if cur.rowcount > 0:
                    self.show_alert("Success", f"Article {article} removed from DNO.")
                    self.new_found_dnos += 1
                else:
                    self.show_alert("Not Found", f"Article {article} was not in DNO.")
            except sqlite3.Error as e:
                self.show_alert("Error", str(e))
            finally:
                conn.close()

    # ------------------------ Inventory in Memory (pandas) ------------------------

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Workbooks", "*.xlsx")])
        if not file_path:
            return

        # 1) Read from Excel
        new_df = pd.read_excel(file_path, engine="openpyxl")

        # 2) Identify recognized departments (first column) -> turn on green lights
        #    You can do a "for row in new_df" approach, or just check unique values in that col.
        departments_in_file = new_df.iloc[:, 0]
        for dep in departments_in_file:
            for outer_key, inner_values in self.departments.items():
                if dep in inner_values:
                    self.lights_bool[outer_key] = True
                    self.lights[outer_key].itemconfig("light", fill="green")
                    break

        # 3) Filter out banned categories
        #    We assume columns: [Department, Merchandise Category, Article Description, Article, Inventory]
        columns = ["Department", "Merchandise Category", "Article Description", "Article", "Inventory"]
        new_df = new_df[columns].copy()

        def is_banned(cat):
            return any(cat.startswith(bad) for bad in self.BANNED_CATS)

        new_df = new_df[~new_df["Merchandise Category"].apply(is_banned)].reset_index(drop=True)

        # 4) If self.df_inventory is empty, just set it; otherwise, append.
        if self.df_inventory.empty:
            self.df_inventory = new_df
        else:
            # Append new data to existing data
            self.df_inventory = pd.concat([self.df_inventory, new_df], ignore_index=True)

        # 5) (Optional) Drop duplicate articles if needed.
        #    Keep the most recent row for the same article, or whichever suits your needs.
        self.df_inventory.drop_duplicates(subset=["Article"], keep="last", inplace=True, ignore_index=True)


    def find_zeros(self):
        if self.df_inventory.empty:
            self.show_alert("No Data", "Please upload Excel first.")
            return

        zero_df = self.df_inventory[self.df_inventory["Inventory"] <= 0].copy()

        # Exclude DNO
        dno_set = self.fetch_dno_articles()

        zero_df = zero_df[~zero_df["Article"].isin(dno_set)]

        for art in zero_df["Article"]:
            self.filtered_zeros.add(art)

        self.zero_count = len(self.filtered_zeros)
        self.update_zero_button()

        self.show_alert("Success", f"Found {self.zero_count} Zeros.\nReady to send to SAP.")

    def find_lows(self):
        if self.df_inventory.empty:
            self.show_alert("No Data", "Please upload Excel first.")
            return

        low_df = self.df_inventory[
            (self.df_inventory["Inventory"] > 0) & (self.df_inventory["Inventory"] <= self.LOW_THRESHOLD)
        ].copy()

        for art in low_df['Article']:
            self.filtered_lows.add(art)

        self.show_alert("Success", f"Found {len(self.filtered_lows)} Lows")
        self.update_low_button(len(self.filtered_lows))

    # ------------------------ Send to SAP (PyAutoGUI) ------------------------

    def send_to_SAP(self, mode=0):
        """
        mode=0 -> send self.filtered_zeros
        mode=1 -> send self.filtered_lows
        """
        entryx, entryy = 222, 330

        def process_lines(data_list):
            if not data_list:
                return
            line = data_list.pop(0).strip()

            if line:
                pyautogui.click(entryx, entryy)
                pyautogui.click(entryx, entryy)
                time.sleep(1.02)
                pyautogui.write(line)
                pyautogui.press("enter")
                time.sleep(0.5)
                pyautogui.press("enter")
                time.sleep(0.5)

            process_lines(data_list)

        if mode == 1:
            # Lows
            if not self.filtered_lows:
                self.show_alert("No Lows", "No low articles to send.")
                return
            confirm = messagebox.askokcancel(
                "Confirm Action",
                f"Sending {len(self.filtered_lows)} lows to SAP.\nMake sure SAP is in the far-left position."
            )
            if confirm:
                time.sleep(2)
                data_list = list(self.filtered_lows)
                process_lines(data_list)
                self.show_alert("Done", "Low-inventory articles sent to SAP.")
        else:
            # Zeros
            if not self.filtered_zeros:
                self.show_alert("No Zeros", "No zero articles to send.")
                return
            confirm = messagebox.askokcancel(
                "Confirm Action",
                f"Sending {len(self.filtered_zeros)} zeros to SAP.\nMake sure SAP is in the far-left position."
            )
            if confirm:
                time.sleep(3)
                data_list = list(self.filtered_zeros)
                process_lines(data_list)
                self.show_alert("Done", "Zero-inventory articles sent to SAP.")

    # ------------------------ Buttons for SAP Submission ------------------------

    def update_zero_button(self):
        if self.zero_button:
            self.zero_button.destroy()
        self.zero_button = tk.Button(
            self.inv_frame,
            text=f"Send {self.zero_count} Zeros to SAP",
            command=self.send_to_SAP
        )
        self.zero_button.pack(padx=5, pady=5, fill=tk.X)

    def update_low_button(self, count):
        if self.low_button:
            self.low_button.destroy()
        self.low_button = tk.Button(
            self.inv_frame,
            text=f"Send {count} Lows to SAP",
            command=lambda: self.send_to_SAP(1)
        )
        self.low_button.pack(padx=5, pady=5, fill=tk.X)

    # ------------------------ Logging & Close ------------------------

    def log_activity(self):
        if self.inputted:
            log_message = (
                f"Session Ended: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Zeros sent: {len(self.filtered_zeros)}\n"
                f"Lows sent: {len(self.filtered_lows)}\n"
                f"{self.new_found_dnos} new DNO articles were tracked.\n"
                "--------------------------------------------\n"
            )
            with open("log.txt", "a") as f:
                f.write(log_message)

    def show_alert(self, title, message):
        """Use this method instead of messagebox.showinfo() directly."""
        messagebox.showinfo(title, message)

    def close_app(self):
        self.log_activity()
        self.root.destroy()


# ------------------------ MAIN ------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("700x450+800+400")
    app = FiltererApp(root)
    root.mainloop()
