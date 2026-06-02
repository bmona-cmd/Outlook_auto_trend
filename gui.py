import tkinter as tk
from tkinter import messagebox
import threading
import os
import platform
from pathlib import Path
 
import scripts.read_mails as mail_reader
 
 
automation_thread = None
 
 
def start_automation():
 
    global automation_thread
 
    if (
        automation_thread
        and automation_thread.is_alive()
    ):
        messagebox.showinfo(
            "Running",
            "Automation is already running."
        )
        return
 
    mail_reader.RUNNING = True
 
    status_label.config(
        text="Status: Running"
    )
 
    automation_thread = threading.Thread(
        target=mail_reader.run_mail_reader,
        daemon=True
    )
 
    automation_thread.start()
 
 
def stop_automation():
 
    mail_reader.RUNNING = False
 
    status_label.config(
        text="Status: Stopped"
    )
 
    messagebox.showinfo(
        "Stopped",
        "Automation stopped."
    )
 
 
def open_excel():
 
    output_folder = Path("output")
 
    if not output_folder.exists():
 
        messagebox.showwarning(
            "Not Found",
            "Output folder does not exist."
        )
 
        return
 
    files = list(
        output_folder.glob("*.xlsx")
    )
 
    if not files:
 
        messagebox.showwarning(
            "Not Found",
            "No Excel files found."
        )
 
        return
 
    latest_file = max(
        files,
        key=lambda f: f.stat().st_mtime
    )
 
    try:
 
        if platform.system() == "Darwin":
 
            os.system(
                f'open "{latest_file}"'
            )
 
        elif platform.system() == "Windows":
 
            os.startfile(
                str(latest_file)
            )
 
    except Exception as e:
 
        messagebox.showerror(
            "Error",
            str(e)
        )
 
 
# ==========================
# GUI
# ==========================
 
root = tk.Tk()
 
root.title(
    "Weekend Mail Automation"
)
 
root.geometry("550x400")
 
root.resizable(False, False)
 
 
title = tk.Label(
    root,
    text="Weekend Mail Automation",
    font=("Arial", 18, "bold")
)
 
title.pack(pady=20)
 
 
status_label = tk.Label(
    root,
    text="Status: Stopped",
    font=("Arial", 12)
)
 
status_label.pack(pady=10)
 
 
start_btn = tk.Button(
    root,
    text="Start Automation",
    width=25,
    height=2,
    command=start_automation
)
 
start_btn.pack(pady=10)
 
 
stop_btn = tk.Button(
    root,
    text="Stop Automation",
    width=25,
    height=2,
    command=stop_automation
)
 
stop_btn.pack(pady=10)
 
 
excel_btn = tk.Button(
    root,
    text="Open Latest Excel Sheet",
    width=25,
    height=2,
    command=open_excel
)
 
excel_btn.pack(pady=10)
 
 
footer = tk.Label(
    root,
    text="Outlook Weekend Monitoring Tool",
    font=("Arial", 9),
    fg="gray"
)
 
footer.pack(
    side="bottom",
    pady=15
)
 
 
root.mainloop()