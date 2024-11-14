import win32com.client
import datetime
import os
import tkinter as tk
import json
import time
import ctypes
from tkinter import messagebox, ttk, PhotoImage
from ttkbootstrap import Style

baserdir = os.path.dirname(os.path.abspath(__file__))

# Configuration file path
CONFIG_FILE = "_internal\\config\\user_config.json"
CONFIG_FILE_MULTI_DAY = "_internal\\config\\user_config_multi_day.json"

def create_config_files_dir():
    if not os.path.exists:
        os.makedirs("_internal\\config", exist_ok=True)
    else:
        pass
    
create_config_files_dir()

# Load all configurations from the JSON file
def load_config_single_day():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def load_config_multi_day():
    if os.path.exists(CONFIG_FILE_MULTI_DAY):
        with open(CONFIG_FILE_MULTI_DAY, "r") as f:
            return json.load(f)
    return {}

# Save configurations to the JSON file
def save_config_single_day(config_data):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config_data, f, indent=4)

# Save configurations to the JSON file
def save_config_multi_day(config_data):
    with open(CONFIG_FILE_MULTI_DAY, "w") as f:
        json.dump(config_data, f, indent=4)
        
class SingleDayForm:
    
    def __init__(self, full_name: str, card_id: str, department: str, work_time_start: str, email_cc: str):
        """User Information to be filled in the form fields."""
        now = datetime.datetime.now()
        date = now.strftime("%d.%m.%Y")
        time_end = now.strftime("%H:%M")
        
        self.full_name = full_name
        self.card_id = card_id
        self.department = department
        self.work_time_start = work_time_start
        self.email_cc = email_cc
        self.work_time_end = time_end
        self.todays_date = date
        self.reason = "Homeoffice"

    def fill_form(self):
        USERNAME = os.getenv("USERNAME")
        outlook = win32com.client.Dispatch("Outlook.Application")
        oft_path = rf"C:\Users\{USERNAME}\AppData\Roaming\Microsoft\Templates\Korrekturbeleg ZEUS.oft"
        # oft_path = "outlook_template\\Korrekturbeleg ZEUS.oft"
        template = outlook.CreateItemFromTemplate(oft_path)

        template.Subject = "Korrekturbeleg"
        template.To = "some@mail.com"
        template.CC = self.email_cc
        
        try:
            template.UserProperties("Name Mitarbeiter").Value = self.full_name
            template.UserProperties("Kartennummer").Value = self.card_id
            template.UserProperties("Abteilung1").Value = self.department
            template.UserProperties("Datum1").Value = self.todays_date
            template.UserProperties("Von1").Value = self.work_time_start
            template.UserProperties("Bis1").Value = self.work_time_end
            template.UserProperties("Grund1").Value = self.reason
        except AttributeError as e:
            messagebox.showerror("Error", f"Error accessing a field: {e}")
            return False

        template.Display()
        return True

class MultiDayForm:
    def __init__(self, full_name: str, card_id: str, department: str, email_cc: str, list_of_dates: list, list_of_start_time: list, list_of_end_time: list):
        self.full_name = full_name
        self.card_id = card_id
        self.department = department
        self.email_cc = email_cc
        self.list_of_dates = list_of_dates
        self.list_of_start_time = list_of_start_time
        self.list_of_end_time = list_of_end_time
        self.reason = "Homeoffice"
        
    def fill_form_multi_day(self):
        USERNAME = os.getenv("USERNAME")
        outlook = win32com.client.Dispatch("Outlook.Application")
        oft_path = rf"C:\Users\{USERNAME}\AppData\Roaming\Microsoft\Templates\Korrekturbeleg ZEUS.oft"
        template = outlook.CreateItemFromTemplate(oft_path)

        template.Subject = "Korrekturbeleg"
        template.To = "some@mail.com" # some@mail.com
        template.CC = self.email_cc
        
        if len(self.list_of_dates) == len(self.list_of_end_time) and len(self.list_of_start_time):
            for i, (date, start_time , end_time) in enumerate(zip(self.list_of_dates, self.list_of_start_time,self.list_of_end_time), 1):
                try:
                    template.UserProperties("Name Mitarbeiter").Value = self.full_name
                    template.UserProperties("Kartennummer").Value = self.card_id
                    template.UserProperties("Abteilung1").Value = self.department
                    template.UserProperties(f"Datum{i}").Value = date
                    template.UserProperties(f"Von{i}").Value = start_time
                    template.UserProperties(f"Bis{i}").Value = end_time
                    template.UserProperties(f"Grund{i}").Value = self.reason
                except AttributeError as e:
                    messagebox.showerror("Error", f"Error accessing a field: {e}")
                    return False
                
            template.Display()
            return True
        
class FormApp:
    def __init__(self, root, single_day_frame, multi_day_frame):
        self.root = root
        self.single_day_frame = single_day_frame
        self.multi_day_frame = multi_day_frame
        self.root.title("Auto-Korrekturbeleg 1.1")
        self.clock = tk.Label()
        
        icon = PhotoImage(file = "_internal\\icon\\working-time.png")
        self.root.iconphoto(False, icon)
        
        # Set modern theme
        self.style = Style(theme="light")
        
        # Original Theme Color
        # self.style = ttk.Style(self.root)
        # self.style.theme_use("clam")

        # Create footer frame for clock
        self.footer_frame = ttk.Frame(root)
        self.footer_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        
        # Load all configurations from the config file
        self.config_single_day = load_config_single_day()
        self.config_multi_day = load_config_multi_day()
        self.tick()
        
        # GUI elements for Single Day tab
        single_day_frame.grid_columnconfigure(1, weight=1)
        single_day_frame.grid_rowconfigure(0, weight=1)
        single_day_frame.grid_rowconfigure(1, weight=1)
        single_day_frame.grid_rowconfigure(2, weight=1)
        single_day_frame.grid_rowconfigure(3, weight=1)
        single_day_frame.grid_rowconfigure(4, weight=1)
        single_day_frame.grid_rowconfigure(5, weight=1)
        single_day_frame.grid_rowconfigure(6, weight=1)
        single_day_frame.grid_rowconfigure(7, weight=1)
        single_day_frame.grid_rowconfigure(8, weight=1)

        # GUI elements for Multi Day tab
        #multi_day_frame.grid_columnconfigure(0, weight=1)
        multi_day_frame.grid_columnconfigure(1, weight=1)
        multi_day_frame.grid_columnconfigure(2, weight=1)
        multi_day_frame.grid_columnconfigure(3, weight=1)
        multi_day_frame.grid_columnconfigure(4, weight=1)
        multi_day_frame.grid_rowconfigure(0, weight=1)
        multi_day_frame.grid_rowconfigure(1, weight=1)
        multi_day_frame.grid_rowconfigure(2, weight=1)
        multi_day_frame.grid_rowconfigure(3, weight=1)
        multi_day_frame.grid_rowconfigure(4, weight=1)
        multi_day_frame.grid_rowconfigure(5, weight=1)
        multi_day_frame.grid_rowconfigure(6, weight=1)
        multi_day_frame.grid_rowconfigure(7, weight=1)
        multi_day_frame.grid_rowconfigure(8, weight=1)
        multi_day_frame.grid_rowconfigure(9, weight=1)
        multi_day_frame.grid_rowconfigure(10, weight=1)
        boldStyle = ttk.Style()
        boldStyle.configure("Bold.TButton", font = ('6','bold'))
        
        # GUI elements for Multi Day
        ttk.Label(multi_day_frame, text="Full Name:").grid(row=0, column=0, padx=10, pady=5, sticky="E")
        self.full_name_entry_multi_day = ttk.Entry(multi_day_frame)
        self.full_name_entry_multi_day.grid(row=0, column=1, padx=10, pady=5, sticky="EW")

        ttk.Label(multi_day_frame, text="Card ID:").grid(row=1, column=0, padx=10, pady=5, sticky="E")
        self.card_id_entry_multi_day = ttk.Entry(multi_day_frame)
        self.card_id_entry_multi_day.grid(row=1, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(multi_day_frame, text="Department:").grid(row=2, column=0, padx=10, pady=5, sticky="E")
        self.department_entry_multi_day = ttk.Entry(multi_day_frame)
        self.department_entry_multi_day.grid(row=2, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(multi_day_frame, text="E-Mail CC:").grid(row=3, column=0, padx=10, pady=5, sticky="E")
        self.email_cc_entry_multi_day = ttk.Entry(multi_day_frame)
        self.email_cc_entry_multi_day.grid(row=3, column=1, padx=10, pady=5, sticky="EW")
        
        # Dropdown for loading configurations
        ttk.Label(multi_day_frame, text="Select Config:").grid(row=4, column=0, padx=10, pady=5, sticky="E")
        self.config_var_multi_day= tk.StringVar(multi_day_frame)
        self.config_dropdown_multi_day = ttk.Combobox(multi_day_frame, textvariable=self.config_var_multi_day, values=list(self.config_multi_day.keys()), state="readonly")
        self.config_dropdown_multi_day.grid(row=4, column=1, padx=10, pady=5, sticky="EW")

        # Buttons
        self.load_config_button_multi_day = ttk.Button(multi_day_frame, text="Load Config", style="Bold.TButton", command=lambda: self.load_config("MultiDay"))
        self.load_config_button_multi_day.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="EW")
        
        mb_multi_day = ttk.Menubutton(multi_day_frame, text="Config file options...", style="primary.Outline.TMenubutton")
        mb_multi_day.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="EW") 
            
        menu_mb_multi_day = tk.Menu(mb_multi_day, tearoff=False)

        menu_mb_multi_day.add_command(label="Save Config", command=self.save_config)
        menu_mb_multi_day.add_command(label="Delete Config", command=self.delete_config_multi_day)
            
        # associate menu with menubutton
        mb_multi_day["menu"] = menu_mb_multi_day
        
        self.clear_fields_multi_day = ttk.Button(multi_day_frame, text="Clear Fields", style="Bold.TButton", command=lambda: self.helper_clear_fields("MultiDay"))
        self.clear_fields_multi_day.grid(row=6, column=2, columnspan=1, padx=10, pady=10, sticky="EW")
        
        # Add Day to Config
        self.add_day_button = ttk.Button(multi_day_frame, text="Add Today's Day", style="Bold.TButton", command=self.add_day)
        self.add_day_button.grid(row=6, column=3, columnspan=1, padx=10, pady=10, sticky="EW")
        
        self.submit_form_button_multi_day = ttk.Button(multi_day_frame, text="Submit Form", style="Bold.TButton", command=self.submit_form_multi_day)
        self.submit_form_button_multi_day.grid(row=6, column=4, columnspan=1, padx=10, pady=10, sticky="EW")
        
        # ===== Workday Dates Label and Entries ====== #

        self.datum_label = ttk.Label(multi_day_frame, text="Workday Dates:", font=("roboto", 10, "bold"))
        self.datum_label.grid(row=0, column=2, padx=5, pady=5, sticky="EW")
        
        self.datum1_entry = ttk.Entry(multi_day_frame)
        self.datum1_entry.grid(row=1, column=2, padx=10, pady=5, sticky="EW")

        self.datum2_entry = ttk.Entry(multi_day_frame)
        self.datum2_entry.grid(row=2, column=2, padx=10, pady=5, sticky="EW")

        self.datum3_entry = ttk.Entry(multi_day_frame)
        self.datum3_entry.grid(row=3, column=2, padx=10, pady=5, sticky="EW")

        self.datum4_entry = ttk.Entry(multi_day_frame)
        self.datum4_entry.grid(row=4, column=2, padx=10, pady=5, sticky="EW")

        self.datum5_entry = ttk.Entry(multi_day_frame)
        self.datum5_entry.grid(row=5, column=2, padx=10, pady=5, sticky="EW")
        
        # ===== Work Start Time Label and Entries ====== #
        
        self.time_start_label = ttk.Label(multi_day_frame, text="Work Start Time:", font=("roboto", 10, "bold"))
        self.time_start_label.grid(row=0, column=3, padx=5, pady=5, sticky="EW")
        
        self.time_work_start1_entry = ttk.Entry(multi_day_frame)
        self.time_work_start1_entry.grid(row=1, column=3, padx=10, pady=5, sticky="EW")
        
        self.time_work_start2_entry = ttk.Entry(multi_day_frame)
        self.time_work_start2_entry.grid(row=2, column=3, padx=10, pady=5, sticky="EW")
        
        self.time_work_start3_entry = ttk.Entry(multi_day_frame)
        self.time_work_start3_entry.grid(row=3, column=3, padx=10, pady=5, sticky="EW")
        
        self.time_work_start4_entry = ttk.Entry(multi_day_frame)
        self.time_work_start4_entry.grid(row=4, column=3, padx=10, pady=5, sticky="EW")
        
        self.time_work_start5_entry = ttk.Entry(multi_day_frame)
        self.time_work_start5_entry.grid(row=5, column=3, padx=10, pady=5, sticky="EW")
        
        # ===== Work End Time Label and Entries ====== #
        
        self.time_from_label = ttk.Label(multi_day_frame, text="Work End Time:", font=("roboto", 10, "bold"))
        self.time_from_label.grid(row=0, column=4, padx=5, pady=5, sticky="EW")
        
        self.time_work_end1_entry = ttk.Entry(multi_day_frame)
        self.time_work_end1_entry.grid(row=1, column=4, padx=10, pady=5, sticky="EW")
        
        self.time_work_end2_entry = ttk.Entry(multi_day_frame)
        self.time_work_end2_entry.grid(row=2, column=4, padx=10, pady=5, sticky="EW")
        
        self.time_work_end3_entry = ttk.Entry(multi_day_frame)
        self.time_work_end3_entry.grid(row=3, column=4, padx=10, pady=5, sticky="EW")
        
        self.time_work_end4_entry = ttk.Entry(multi_day_frame)
        self.time_work_end4_entry.grid(row=4, column=4, padx=10, pady=5, sticky="EW")
        
        self.time_work_end5_entry = ttk.Entry(multi_day_frame)
        self.time_work_end5_entry.grid(row=5, column=4, padx=10, pady=5, sticky="EW")
        
        # ================================================================================================ #
        
        # GUI elements for Single Day
        ttk.Label(single_day_frame, text="Full Name:").grid(row=0, column=0, padx=10, pady=5, sticky="E")
        self.full_name_entry = ttk.Entry(single_day_frame)
        self.full_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="EW")

        ttk.Label(single_day_frame, text="Card ID:").grid(row=1, column=0, padx=10, pady=5, sticky="E")
        self.card_id_entry = ttk.Entry(single_day_frame)
        self.card_id_entry.grid(row=1, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(single_day_frame, text="Department:").grid(row=2, column=0, padx=10, pady=5, sticky="E")
        self.department_entry = ttk.Entry(single_day_frame)
        self.department_entry.grid(row=2, column=1, padx=10, pady=5, sticky="EW")

        ttk.Label(single_day_frame, text="Work Time Start:").grid(row=3, column=0, padx=10, pady=5, sticky="E")
        self.work_time_start_entry = ttk.Entry(single_day_frame)
        self.work_time_start_entry.grid(row=3, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(single_day_frame, text="E-Mail CC:").grid(row=4, column=0, padx=10, pady=5, sticky="E")
        self.email_cc_entry = ttk.Entry(single_day_frame)
        self.email_cc_entry.grid(row=4, column=1, padx=10, pady=5, sticky="EW")
        
        # Dropdown for loading configurations
        ttk.Label(single_day_frame, text="Select Config:").grid(row=5, column=0, padx=10, pady=5, sticky="E")
        self.config_var = tk.StringVar(single_day_frame)
        self.config_dropdown = ttk.Combobox(single_day_frame, textvariable=self.config_var, values=list(self.config_single_day.keys()), state="readonly")
        self.config_dropdown.grid(row=5, column=1, padx=10, pady=5, sticky="EW")
        
        # Clock
        self.clock = ttk.Label(self.footer_frame, font=("roboto", 16, "bold"), background="#dcdad5")
        self.clock.grid(row=8, columnspan=2, padx=10, pady=10, sticky="EW")
        self.clock.pack(expand=True)
        
        mb = ttk.Menubutton(single_day_frame, text="Config file options...", style="primary.Outline.TMenubutton")
        mb.grid(row=6, column=1, padx=10, pady=10, sticky="EW")  # grid(row=9, columnspan=2, padx=10, pady=10, sticky="EW")
            
        menu = tk.Menu(mb, tearoff=False)

        menu.add_command(label="Save Config", command=self.save_config)
        menu.add_command(label="Delete Config", command=self.delete_config)
            
        # associate menu with menubutton
        mb["menu"] = menu
        
        # Buttons
        self.load_config_button = ttk.Button(single_day_frame, text="Load Config", style="Bold.TButton", command=lambda: self.load_config("SingleDay"))
        self.load_config_button.grid(row=6, column=0, padx=10, pady=10, sticky="EW")
        self.clear_fields_button = ttk.Button(single_day_frame, text="Clear Fields", style="Bold.TButton", command=lambda:self.helper_clear_fields("SingleDay"))
        self.clear_fields_button.grid(row=7, column=0, padx=10, pady=10, sticky="EW")
        self.submit_form_button = ttk.Button(single_day_frame, text="Submit Form", style="Bold.TButton", command=self.submit_form)
        self.submit_form_button.grid(row=7, column=1, columnspan=1, padx=10, pady=10, sticky="EW")
        

    def get_work_start_time(self):
        """This method takes the “System Up Time” from Task Manager and converts the time into seconds.\n
        It then gets the current system time in seconds and subtracts the System Up Time from it to calculate the start time.\n
        The start time is then converted into the format hour:minute.

        Returns:
            int: hour, minute
        """

        # Get the library where GetTickCount64 resides
        lib = ctypes.windll.kernel32

        # Get system uptime in milliseconds
        t = lib.GetTickCount64()

        # Convert milliseconds to seconds
        sys_uptime_seconds = t // 1000

        # Get current time
        now = datetime.datetime.now()

        # Calculate the start time by subtracting the uptime from the current time
        work_start_time = now - datetime.timedelta(seconds=sys_uptime_seconds)

        # Extract hours and minutes from the start time
        return work_start_time.hour, work_start_time.minute


    def load_config(self, parent):
        selected_config = self.config_var.get()
        selected_config_multi_day = self.config_var_multi_day.get()
        
        # Single Day Config
        if parent ==  "SingleDay":
            if selected_config and selected_config in self.config_single_day:
                config_data = self.config_single_day[selected_config]
                self.full_name_entry.delete(0, tk.END)
                self.full_name_entry.insert(0, config_data.get("full_name", ""))
                self.card_id_entry.delete(0, tk.END)
                self.card_id_entry.insert(0, config_data.get("card_id", ""))
                self.department_entry.delete(0, tk.END)
                self.department_entry.insert(0, config_data.get("department", ""))
                self.work_time_start_entry.delete(0, tk.END)
                self.work_time_start_entry.insert(0, config_data.get("work_time_start", ""))
                self.email_cc_entry.delete(0, tk.END)
                self.email_cc_entry.insert(0, config_data.get("email_cc",""))
            else:
                messagebox.showerror("No config file selected", "No config file to load, please select one from the menu.")
            
        # Multi Day Config
        if parent ==  "MultiDay":
            if selected_config_multi_day and selected_config_multi_day in self.config_multi_day:
                config_data_multi_day = self.config_multi_day[selected_config_multi_day]
                self.full_name_entry_multi_day.delete(0, tk.END)
                self.full_name_entry_multi_day.insert(0, config_data_multi_day.get("full_name", ""))
                self.card_id_entry_multi_day.delete(0, tk.END)
                self.card_id_entry_multi_day.insert(0, config_data_multi_day.get("card_id", ""))
                self.department_entry_multi_day.delete(0, tk.END)
                self.department_entry_multi_day.insert(0, config_data_multi_day.get("department", ""))
                self.email_cc_entry_multi_day.delete(0, tk.END)
                self.email_cc_entry_multi_day.insert(0, config_data_multi_day.get("email_cc",""))
                self.datum1_entry.delete(0, tk.END)
                self.datum1_entry.insert(0, config_data_multi_day.get("weekday1", ""))
                self.datum2_entry.delete(0, tk.END)
                self.datum2_entry.insert(0, config_data_multi_day.get("weekday2", ""))
                self.datum3_entry.delete(0, tk.END)
                self.datum3_entry.insert(0, config_data_multi_day.get("weekday3", ""))
                self.datum4_entry.delete(0, tk.END)
                self.datum4_entry.insert(0, config_data_multi_day.get("weekday4", ""))
                self.datum5_entry.delete(0, tk.END)
                self.datum5_entry.insert(0, config_data_multi_day.get("weekday5", ""))
                self.time_work_start1_entry.delete(0, tk.END)
                self.time_work_start1_entry.insert(0, config_data_multi_day.get("starttime1", ""))
                self.time_work_start2_entry.delete(0, tk.END)
                self.time_work_start2_entry.insert(0, config_data_multi_day.get("starttime2", ""))
                self.time_work_start3_entry.delete(0, tk.END)
                self.time_work_start3_entry.insert(0, config_data_multi_day.get("starttime3", ""))
                self.time_work_start4_entry.delete(0, tk.END)
                self.time_work_start4_entry.insert(0, config_data_multi_day.get("starttime4", ""))
                self.time_work_start5_entry.delete(0, tk.END)
                self.time_work_start5_entry.insert(0, config_data_multi_day.get("endtime5", ""))
                self.time_work_end1_entry.delete(0, tk.END)
                self.time_work_end1_entry.insert(0, config_data_multi_day.get("endtime1", ""))
                self.time_work_end2_entry.delete(0, tk.END)
                self.time_work_end2_entry.insert(0, config_data_multi_day.get("endtime2", ""))
                self.time_work_end3_entry.delete(0, tk.END)
                self.time_work_end3_entry.insert(0, config_data_multi_day.get("endtime3", ""))
                self.time_work_end4_entry.delete(0, tk.END)
                self.time_work_end4_entry.insert(0, config_data_multi_day.get("endtime4", ""))
                self.time_work_end5_entry.delete(0, tk.END)
                self.time_work_end5_entry.insert(0, config_data_multi_day.get("endtime5", ""))
            else:
                messagebox.showerror("No config file selected", "No config file to load, please select one from the menu.")
    
    def save_config(self):
        config_name = self.full_name_entry.get()
        config_name_multi_day = self.full_name_entry_multi_day.get()
        
        if config_name:
            new_config = {
                "full_name": self.full_name_entry.get(),
                "card_id": self.card_id_entry.get(),
                "department": self.department_entry.get(),
                "work_time_start": self.work_time_start_entry.get(),
                "email_cc": self.email_cc_entry.get()
            }
            self.config_single_day[config_name] = new_config
            save_config_single_day(self.config_single_day)
            self.config_dropdown["values"] = list(self.config_single_day.keys())
            
        if config_name_multi_day:
            new_config_multi_day = {
                "full_name": self.full_name_entry_multi_day.get(),
                "card_id": self.card_id_entry_multi_day.get(),
                "department": self.department_entry_multi_day.get(),
                "email_cc": self.email_cc_entry_multi_day.get(),
                "weekday1": self.datum1_entry.get(),
                "weekday2": self.datum2_entry.get(),
                "weekday3": self.datum3_entry.get(),
                "weekday4": self.datum4_entry.get(),
                "weekday5": self.datum5_entry.get(),
                "starttime1": self.time_work_start1_entry.get(),
                "starttime2": self.time_work_start2_entry.get(),
                "starttime3": self.time_work_start3_entry.get(),
                "starttime4": self.time_work_start4_entry.get(),
                "starttime5": self.time_work_start5_entry.get(),
                "endtime1": self.time_work_end1_entry.get(),
                "endtime2": self.time_work_end2_entry.get(),
                "endtime3": self.time_work_end3_entry.get(),
                "endtime4": self.time_work_end4_entry.get(),
                "endtime5": self.time_work_end5_entry.get()
            }
            self.config_multi_day[config_name_multi_day] = new_config_multi_day
            save_config_multi_day(self.config_multi_day)
            self.config_dropdown_multi_day["values"] = list(self.config_multi_day.keys())

        if config_name or config_name_multi_day:
            messagebox.showinfo("Success", "Configuration saved!")
        else:
            messagebox.showerror("Error", "Full Name is required to save config.")


    def delete_config(self):
        config_entry = self.config_dropdown.get()
        current_values = self.config_dropdown["values"] = list(self.config_single_day.keys())
        parent = "SingleDay"
        if config_entry:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r") as f:
                    json_data = json.load(f)

                if config_entry in json_data:
                    json_data.pop(config_entry)
                    if config_entry in current_values:
                        self.config_single_day.pop(config_entry)
                        current_values.remove(config_entry)
                        self.config_dropdown["values"] = current_values 
                        self.config_dropdown.set("")

            # saving the updated JSON data back to the file
            with open(CONFIG_FILE, "w") as f:
                json.dump(json_data, f, indent=2)
                
            self.helper_clear_fields(parent)
        else:
            messagebox.showerror("File not found", "No config file found to delete, please select one from the dropdown menu.")
    
    
    def delete_config_multi_day(self):
        config_entry_multi_day = self.config_dropdown_multi_day.get()
        current_values_multi_day = self.config_dropdown_multi_day["values"] = list(self.config_multi_day.keys())
        parent = "MultiDay"
        
        if config_entry_multi_day:
            if os.path.exists(CONFIG_FILE_MULTI_DAY):
                with open(CONFIG_FILE_MULTI_DAY) as f:
                    json_data = json.load(f)
                
                if config_entry_multi_day in json_data:
                    json_data.pop(config_entry_multi_day)
                    if config_entry_multi_day in current_values_multi_day:
                        self.config_multi_day.pop(config_entry_multi_day)
                        current_values_multi_day.remove(config_entry_multi_day)
                        self.config_dropdown_multi_day["values"] = current_values_multi_day
                        self.config_dropdown_multi_day.set("")
                        
            # saving the updated JSON data back to the file
            with open(CONFIG_FILE_MULTI_DAY, "w") as f:
                json.dump(json_data, f, indent=2)
            
            self.helper_clear_fields(parent)
        else:
            messagebox.showerror("File not found", "No config file found to delete, please select one from the dropdown menu.")


    def add_day(self):
        now = datetime.datetime.now()
        todays_date = now.strftime("%d.%m.%Y")
        time_end = now.strftime("%H:%M")
        date_entries = [self.datum1_entry, self.datum2_entry, self.datum3_entry, self.datum4_entry, self.datum5_entry]
        end_time_entries = [self.time_work_end1_entry, self.time_work_end2_entry, self.time_work_end3_entry, self.time_work_end4_entry, self.time_work_end5_entry]
        start_time_entries = [self.time_work_start1_entry, self.time_work_start2_entry, self.time_work_start3_entry, self.time_work_start4_entry, self.time_work_start5_entry]
        try:
            # Check if today's date already exists in any entry
            if todays_date in [entry.get() for entry in date_entries]:
                messagebox.showerror("Error", f"Date '{todays_date}' already exists in entries.")
                return

            # Find first empty date slot and corresponding time slot
            for date_entry, end_time_entry, start_time_entry in zip(date_entries, end_time_entries, start_time_entries):
                if len(date_entry.get()) == 0:
                    date_entry.insert(0, todays_date)
                    hour, minute = self.get_work_start_time()
                    if len(end_time_entry.get()) == 0:
                        end_time_entry.insert(0, time_end)
                    if len(start_time_entry.get()) == 0:
                        start_time_entry.insert(0, f"{hour}:{minute}")
                    return  # Exit after successful insertion
                    
            # If we get here, all slots are full
            messagebox.showerror("Error", "All date slots are filled. Cannot add more entries.")
                
        except Exception as ex:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(ex)}")
    
    
    def submit_form_multi_day(self):
        full_name = self.full_name_entry_multi_day.get()
        card_id = self.card_id_entry_multi_day.get()
        department = self.department_entry_multi_day.get()
        email_cc = self.email_cc_entry_multi_day.get()
        
        date_entries = [entry for entry in [self.datum1_entry.get(), self.datum2_entry.get(), 
                                    self.datum3_entry.get(), self.datum4_entry.get(), 
                                    self.datum5_entry.get()] if entry != ""]
        
        start_time_entries = [entry for entry in [self.time_work_start1_entry.get(), self.time_work_start2_entry.get(), 
                                                self.time_work_start3_entry.get(), self.time_work_start4_entry.get(), self.time_work_start5_entry.get()] if entry != ""]

        end_time_entries = [entry for entry in [self.time_work_end1_entry.get(), self.time_work_end2_entry.get(), 
                            self.time_work_end3_entry.get(), self.time_work_end4_entry.get(), self.time_work_end5_entry.get()] if entry != ""]
        
        if not full_name or not card_id or not department  or not email_cc or not date_entries or not start_time_entries or not end_time_entries:
            messagebox.showerror("Input Error", "All fields are required! Also at least one date, one start and one end time field!")
            return
        
        user_info = MultiDayForm(full_name, card_id, department, email_cc, date_entries, start_time_entries, end_time_entries)
        if user_info.fill_form_multi_day():
            messagebox.showinfo("Success", "Form submitted successfully!")
        else:
            messagebox.showerror("Error", "An error occurred while trying to submit the form.")
    
    
    def submit_form(self):
        full_name = self.full_name_entry.get()
        card_id = self.card_id_entry.get()
        department = self.department_entry.get()
        work_time_start = self.work_time_start_entry.get()
        email_cc = self.email_cc_entry.get()

        if not full_name or not card_id or not department or not work_time_start or not email_cc:
            messagebox.showerror("Input Error", "All fields are required!")
            return
        
        user_info = SingleDayForm(full_name, card_id, department, work_time_start, email_cc)
        if user_info.fill_form():
            messagebox.showinfo("Success", "Form submitted successfully!")
        else:
            messagebox.showerror("Error", "An error occurred while trying to submit the form.")
            
            
    def tick(self):
        time_string = time.strftime("%H:%M:%S")
        date_string = time.strftime("%d.%m.%Y")
        self.clock.config(text=time_string + " / " + date_string)
        self.clock.after(1000, self.tick)
    
    
    def refresh(self):
        root.update()
        
    
    def helper_clear_fields(self, parent):

        single_day_entires = [item for item in [self.full_name_entry, self.card_id_entry,
                                self.department_entry,self.work_time_start_entry, self.email_cc_entry] if item != ""]
        
        multi_day_entries = [item for item in [self.datum1_entry,
                                            self.datum2_entry, self.datum3_entry, 
                                            self.datum4_entry, self.datum4_entry,
                                            self.time_work_start1_entry, self.time_work_start2_entry,
                                            self.time_work_start3_entry, self.time_work_start4_entry,
                                            self.time_work_start5_entry,
                                            self.time_work_end1_entry, self.time_work_end2_entry,
                                            self.time_work_end3_entry, self.time_work_end4_entry, 
                                            self.time_work_end5_entry] if item != ""]
        
        if parent == "SingleDay":
            for entry in single_day_entires:
                entry.delete(0, tk.END)
        
        if parent == "MultiDay":
            for entry in multi_day_entries:
                entry.delete(0, tk.END)


    # Function to resize the window based on the selected tab

def resize_window(event):
    selected_tab = notebook.index(notebook.select())
    if selected_tab == 0:  # First tab
        root.geometry("350x400")    # Set the desired width and height for the first tab
    elif selected_tab == 1:  # Second tab
        root.geometry("650x400")  # Set the desired width and height for the second tab

def on_closing():
    if messagebox.askyesno("Quit", "Do you want to quit?"):
        root.destroy()
    
# Create the Tkinter root window
root = tk.Tk()
root.geometry("350x400+1100+500")
root.wm_resizable(False, False)
root.protocol("WM_DELETE_WINDOW", on_closing)

# Notebook for tabs
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Main Tab
single_day_tab = ttk.Frame(notebook)
notebook.add(single_day_tab, text="Single Day")

# Multi-Day Tab
multi_day_tab = ttk.Frame(notebook)
notebook.add(multi_day_tab, text="Multiple Days")

# Bind the NotebookTabChanged
notebook.bind("<<NotebookTabChanged>>", resize_window)

# Initialize the form app
app = FormApp(root, single_day_tab, multi_day_tab)

# Start the Tkinter event loop
root.mainloop()
