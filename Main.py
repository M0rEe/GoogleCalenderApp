import tkinter as tk
from tkinter import filedialog, messagebox, ttk
##########################################
import pandas as pd 
import re
##########################################
from gcsa.google_calendar import GoogleCalendar
from gcsa.event import Event
import datetime as dt
##########################################


def get_next_weekday(day_name:str)-> dt:
    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    today = dt.datetime.now()
    today_day_index = today.weekday()  # Monday is 0 and Sunday is 6
    target_day_index = days_of_week.index(day_name)
    days_ahead = (target_day_index - today_day_index + 7) % 7
    if days_ahead == 0:
        days_ahead = 7  # If it's the same day, get the next week's day
    next_day = today + dt.timedelta(days=days_ahead)
    return next_day.date()


def execute_data_checking(email:str)-> GoogleCalendar:
    try:
        calendar = GoogleCalendar(email, credentials_path="./credentials/credentials.json")
        return calendar
    except Exception as e:
        messagebox.showerror("Error", f"Failed to connect to Google Calendar. Error: {e}")
        return None

def read_execl_sheet_data(path:str)-> list:
    try:
        df = pd.read_excel(path)
        dropped_columns = ['user_doc_id', 'modified_at_iso', 'status', 'review_url', 'uploaded_from', 'doc_meta_data',
                           'folder_name', 'folder_id', 'title', 'doc_id', 'created_at_iso', 'type']
        df = df.drop(columns=dropped_columns, axis=1)
        df.columns = df.columns.str.replace('Tables ', '', regex=True)
        row = 0
        lst = [[]]

        for x in range(df.shape[0] - 2):
            for y in range(df.shape[1]):
                value = df.iloc[x, y]
                if not pd.isna(value):
                    lst[row].append(df.columns[y])
                    if isinstance(value, str) and re.match(r'[0-9]', str(value)):
                        temp = value.split('-')
                        lst[row].append(temp[0])
                        lst[row].append(temp[1])
                    else:
                        lst[row].append(value)
            lst.append([])
            row += 1

        lst.pop()
        return lst
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file. Error: {e}")
        return []
    
def add_entries_to_calendar(lst:list , calendar:GoogleCalendar,progress_var:int) -> None:
    try:
        for i, x in enumerate(lst):
            if len(x) > 4:
                date_value = get_next_weekday(x[2])
                start_hour = x[3].split(":")
                end_hour = x[4].split(":")
                event = Event(
                    x[1],
                    start=dt.datetime(date_value.year, date_value.month, date_value.day, int(start_hour[0]) + 12,
                                      int(start_hour[1]) + 12 if len(start_hour) > 1 else 0),
                    location=x[6],
                    end=dt.datetime(date_value.year, date_value.month, date_value.day, int(end_hour[0]) + 12,
                                    int(end_hour[1]) + 12 if len(end_hour) > 1 else 0),
                    minutes_before_popup_reminder=30,
                    minutes_before_email_reminder=90
                )
                calendar.add_event(event)

            # Update progress bar
            progress_var.set((i + 1) / len(lst) * 100)
            root.update_idletasks()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to add events to calendar. Error: {e}")



def start_conversion():
    email = email_entry.get()
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        messagebox.showerror("Invalid Email", "Please enter a valid email address.")
        return

    calendar = execute_data_checking(email)
    if not calendar:
        return

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        lst = read_execl_sheet_data(file_path)
        if lst:
            add_entries_to_calendar(lst, calendar, progress_var)
            
            
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Add Events from Excel to Google Calendar")
    root.geometry("600x300")

    label = tk.Label(root, text="Enter your Google Calendar email:", font=("Arial", 12))
    label.pack(pady=10)

    # Email entry with validation
    email_entry = tk.Entry(root, font=("Arial", 16),width=40)
    email_entry.pack(pady=10)

    # Create a button to open the file dialog and start the conversion
    convert_button = tk.Button(root, text="Select Excel File & Start", font=("Arial", 12), command=start_conversion)
    convert_button.pack(pady=10)

    # Progress bar
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(pady=20, fill=tk.X)

    root.mainloop()

