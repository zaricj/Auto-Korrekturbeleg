import win32com.client
import datetime
import os
import tkinter as tk
from tkinter import messagebox, ttk
import json
import time

baserdir = os.path.dirname(__file__)

# Configuration file path
CONFIG_FILE = f"{baserdir}\\config\\user_config.json"

# Load all configurations from the JSON file
def load_all_configs():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

# Save configurations to the JSON file
def save_config(config_data):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config_data, f, indent=4)

class UserInformation:
    
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
        template = outlook.CreateItemFromTemplate(oft_path)

        template.Subject = "Korrekturbeleg"
        template.To = "Zeiterfassung.NES@Geis-Group.de"
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
            print(f"Error accessing a field: {e}")

        template.Display()

class FormApp:
    
    def __init__(self, root, main_frame):
        self.root = root
        self.root.title("Zeiterfassung Korrekturbeleg")
        self.clock = tk.Label()
        
        # Allow resizing
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=3)
        self.root.grid_rowconfigure(6, weight=1)
        
        # Set modern theme
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        
        # Load all configurations from the config file
        self.all_configs = load_all_configs()
        self.tick()
        
        # GUI elements
        ttk.Label(main_frame, text="Full Name:").grid(row=0, column=0, padx=10, pady=5, sticky="E")
        self.full_name_entry = ttk.Entry(main_frame)
        self.full_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="EW")

        ttk.Label(main_frame, text="Card ID:").grid(row=1, column=0, padx=10, pady=5, sticky="E")
        self.card_id_entry = ttk.Entry(main_frame)
        self.card_id_entry.grid(row=1, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(main_frame, text="Department:").grid(row=2, column=0, padx=10, pady=5, sticky="E")
        self.department_entry = ttk.Entry(main_frame)
        self.department_entry.grid(row=2, column=1, padx=10, pady=5, sticky="EW")

        ttk.Label(main_frame, text="Work Time Start:").grid(row=3, column=0, padx=10, pady=5, sticky="E")
        self.work_time_start_entry = ttk.Entry(main_frame)
        self.work_time_start_entry.grid(row=3, column=1, padx=10, pady=5, sticky="EW")
        
        ttk.Label(main_frame, text="E-Mail CC:").grid(row=4, column=0, padx=10, pady=5, sticky="E")
        self.email_cc_entry = ttk.Entry(main_frame)
        self.email_cc_entry.grid(row=4, column=1, padx=10, pady=5, sticky="EW")
        
        # Dropdown for loading configurations
        ttk.Label(main_frame, text="Select Config:").grid(row=5, column=0, padx=10, pady=5, sticky="E")
        self.config_var = tk.StringVar(main_frame)
        self.config_dropdown = ttk.Combobox(main_frame, textvariable=self.config_var, values=list(self.all_configs.keys()), state="readonly")
        self.config_dropdown.grid(row=5, column=1, padx=10, pady=5, sticky="EW")
        
        # Clock
        self.clock = ttk.Label(main_frame, font=("roboto", 16, "bold"), background="#dcdad5")
        self.clock.grid(row=8, columnspan=2, padx=10, pady=10, sticky="EW")
        
        mb = ttk.Menubutton(main_frame, text="Config file options...", style="info.Outline.TMenubutton")
        mb.grid(row=6, column=1, padx=10, pady=10, sticky="EW")  # grid(row=9, columnspan=2, padx=10, pady=10, sticky="EW")
            
        menu = tk.Menu(mb)

        menu.add_checkbutton(label="Save Config", command=self.save_config)
        menu.add_command(label="Delete Config", command=self.delete_config)
            
        # associate menu with menubutton
        mb["menu"] = menu
        
        # Buttons
        ttk.Button(main_frame, text="Load Config", command=self.load_config).grid(row=6, column=0, padx=10, pady=10, sticky="EW")
        # ttk.Button(main_frame, text="Save Config", command=self.save_config).grid(row=6, column=1, padx=10, pady=10, sticky="EW")
        ttk.Button(main_frame, text="Submit", command=self.submit_form).grid(row=7, columnspan=2, padx=10, pady=10, sticky="EW")
        
        # Configure row and column weights for the main tab

        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_rowconfigure(4, weight=1)
        main_frame.grid_rowconfigure(5, weight=1)
        main_frame.grid_rowconfigure(6, weight=1)
        main_frame.grid_rowconfigure(7, weight=1)
        main_frame.grid_rowconfigure(8, weight=1)
        

    def load_config(self):
        selected_config = self.config_var.get()
        if selected_config and selected_config in self.all_configs:
            config_data = self.all_configs[selected_config]
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

    def save_config(self):
        config_name = self.full_name_entry.get()
        if config_name:
            new_config = {
                "full_name": self.full_name_entry.get(),
                "card_id": self.card_id_entry.get(),
                "department": self.department_entry.get(),
                "work_time_start": self.work_time_start_entry.get(),
                "email_cc": self.email_cc_entry.get()
            }
            self.all_configs[config_name] = new_config
            save_config(self.all_configs)
            self.config_dropdown['values'] = list(self.all_configs.keys())
            messagebox.showinfo("Success", "Configuration saved!")
        else:
            messagebox.showerror("Error", "Full Name is required to save config.")

    def delete_config(self):
        config_entry = self.config_dropdown.get()
        current_values = self.config_dropdown['values'] = list(self.all_configs.keys())
        
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                json_data = json.load(f)

            if config_entry in json_data:
                json_data.pop(config_entry)
                if config_entry in current_values:
                    current_values.remove(config_entry)
                    self.config_dropdown['values'] = current_values 
                    self.config_dropdown.set("")

            # saving the updated JSON data back to the file
        with open(CONFIG_FILE, 'w') as f:
            json.dump(json_data, f, indent=2)
            
    
    def submit_form(self):
        full_name = self.full_name_entry.get()
        card_id = self.card_id_entry.get()
        department = self.department_entry.get()
        work_time_start = self.work_time_start_entry.get()
        email_cc = self.email_cc_entry.get()

        if not full_name or not card_id or not department or not work_time_start or not email_cc:
            messagebox.showerror("Input Error", "All fields are required!")
            return
        
        user_info = UserInformation(full_name, card_id, department, work_time_start, email_cc)
        user_info.fill_form()
        messagebox.showinfo("Success", "Form submitted successfully!")
        
    def tick(self):
        time_string = time.strftime("%H:%M:%S")
        date_string = time.strftime("%d.%m.%Y")
        self.clock.config(text=time_string + " / " + date_string)
        self.clock.after(200, self.tick)
    
    def refresh(self):
        root.update()
        
    
# Create the Tkinter root window
root = tk.Tk()
root.geometry("+%d+%d" %(1100,500))

# Notebook for tabs
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Main Tab
main_tab = ttk.Frame(notebook)
notebook.add(main_tab, text="Current Day")

# Multi-Day Tab
multi_day_tab = ttk.Frame(notebook)
notebook.add(multi_day_tab, text="Multi Day")

# Initialize the form app
app = FormApp(root, main_tab)
# Start the Tkinter event loop
root.mainloop()
