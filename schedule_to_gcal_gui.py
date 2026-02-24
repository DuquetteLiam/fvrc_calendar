#!/usr/bin/env python3
import os
import sys
import csv
import re
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import scrolledtext, messagebox, simpledialog

# ----------------------------
# Helper functions
# ----------------------------
def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform.startswith("darwin"):
        os.system(f'open "{path}"')
    else:  # Linux
        os.system(f'xdg-open "{path}"')

def parse_time_range(text):
    """Return start_time, end_time in HH:MM 24h format; end_time can be None"""
    text = text.replace("–", "-")
    time_pattern = r'(\d{1,2}(:\d{2})?)\s*(am|pm)?\s*(?:-\s*(\d{1,2}(:\d{2})?)\s*(am|pm)?)?'
    match = re.match(time_pattern, text.strip(), re.IGNORECASE)
    if not match:
        return None, None
    start, _, start_ampm, end, _, end_ampm = match.groups()
    def to_24h(t, ampm):
        t = t.strip()
        if ':' not in t:
            t += ':00'
        h, m = map(int, t.split(':'))
        if ampm:
            ampm = ampm.lower()
            if ampm == 'pm' and h != 12:
                h += 12
            if ampm == 'am' and h == 12:
                h = 0
        return f"{h:02d}:{m:02d}"
    start24 = to_24h(start, start_ampm)
    end24 = to_24h(end, end_ampm) if end else None
    return start24, end24

# ----------------------------
# Main schedule parser
# ----------------------------
def parse_schedule(text, year):
    lines = text.splitlines()
    events = []
    current_date = None
    start_month = None
    last_day = None

    # Detect month from header line
    header_match = re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', lines[0].strip())
    if header_match:
        start_month = datetime.strptime(header_match.group(1), "%b").month
    else:
        start_month = simpledialog.askinteger("Input Needed", "Could not detect month from header. Enter start month (1-12):")
    
    for line in lines:
        line = line.strip()
        if not line or line.startswith("*"):
            continue

        # Detect day (like '23 Mon', '24 Tues') and capture any trailing events on same line
        day_match = re.match(r'(\d{1,2})\s+\w+(.*)', line)
        if day_match:
            day = int(day_match.group(1))
            trailing = day_match.group(2).strip()
            month = start_month

            # If day < last_day, assume month rollover
            if last_day and day < last_day:
                month += 1
                if month > 12:
                    month = 1
            last_day = day

            # Validate date
            try:
                current_date = datetime(year, month, day)
            except ValueError:
                # Invalid date: roll forward from the 28th of the month
                corrected_date = datetime(year, month, 28) + timedelta(days=1)
                messagebox.showwarning("Date corrected",
                    f"Invalid date {month}/{day}/{year}, using {corrected_date.strftime('%b %d')}")
                current_date = corrected_date

            # If there is trailing text on the same line (events), continue processing it
            if not trailing:
                continue
            # replace line with trailing content so it gets parsed as events below
            line = trailing

        if current_date is None:
            continue

        # Split multiple events in one line
        items = re.split(r'•|\s{2,}', line)
        for item in items:
            item = item.strip()
            if not item:
                continue
            start_time, end_time = parse_time_range(item)
            # Remove the time from the description
            desc = re.sub(r'^(\d{1,2}(:\d{2})?\s*(am|pm)?\s*-?\s*\d{0,2}(:\d{2})?\s*(am|pm)?)\s*', '', item, flags=re.IGNORECASE).strip()
            if not desc:
                desc = item
            events.append({
                "Subject": desc,
                "Start Date": current_date.strftime("%m/%d/%Y"),
                "Start Time": start_time if start_time else "",
                "End Date": current_date.strftime("%m/%d/%Y") if end_time else "",
                "End Time": end_time if end_time else "",
                "All Day Event": "True" if not start_time else "False",
                "Description": "",
                "Location": ""
            })
    return events

# ----------------------------
# Save CSV
# ----------------------------
def save_csv(events):
    home = os.path.expanduser("~")
    folder = os.path.join(home, "Documents", "fvrc_calendar_exports")
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, "fvrc_calendar.csv")
    fieldnames = ["Subject", "Start Date", "Start Time", "End Date", "End Time",
                  "All Day Event", "Description", "Location"]
    with open(file_path, "w", newline='', encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(events)
    messagebox.showinfo("Success", f"CSV saved to:\n{file_path}")
    open_file(file_path)

# ----------------------------
# Preview window
# ----------------------------
def preview_events(events):
    result = {"ok": False}
    preview = tk.Toplevel()
    preview.title("Preview Events")
    # Use a larger default size and a reasonable minimum so buttons are visible
    preview.geometry("800x500")
    preview.minsize(640, 380)

    tk.Label(preview, text="Preview — confirm to generate CSV", font=(None, 11, "bold")).pack(anchor="w", padx=8, pady=(6,0))
    text = scrolledtext.ScrolledText(preview, wrap=tk.WORD)
    text.pack(expand=True, fill=tk.BOTH, padx=8, pady=(4,0))
    for e in events:
        start_time = e["Start Time"] if e["Start Time"] else ""
        end_time = e["End Time"] if e["End Time"] else ""
        line = f'{e["Start Date"]} {start_time}-{end_time} {e["Subject"]}\n'
        text.insert(tk.END, line)
    text.config(state=tk.DISABLED)

    def on_generate():
        result["ok"] = True
        preview.destroy()

    def on_back():
        preview.destroy()

    btn_frame = tk.Frame(preview)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=8)
    tk.Button(btn_frame, text="Generate CSV", width=14, bg="green", fg="white", font=(None, 10, "bold"), command=on_generate).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_frame, text="Go back", width=10, command=on_back).pack(side=tk.LEFT, padx=8)

    try:
        preview.grab_set()
    except Exception:
        pass

    # Make sure the dialog is visible and focused so the buttons render
    try:
        preview.lift()
        preview.attributes("-topmost", True)
        preview.update_idletasks()
        preview.focus_force()
        # remove topmost after a short delay so it behaves normally afterwards
        preview.after(200, lambda: preview.attributes("-topmost", False))
    except Exception:
        pass

    preview.wait_window()
    return result["ok"]

# ----------------------------
# GUI
# ----------------------------
def generate_csv():
    text = text_input.get("1.0", tk.END)
    year = year_input.get()
    if not year.isdigit():
        messagebox.showerror("Error", "Year must be a number")
        return
    events = parse_schedule(text, int(year))
    if not events:
        messagebox.showerror("Error", "No events found")
        return
    ok = preview_events(events)
    if not ok:
        return
    save_csv(events)

root = tk.Tk()
root.title("FVRC Schedule to Google Calendar CSV")
root.geometry("700x500")

tk.Label(root, text="Paste schedule text here:").pack(anchor="w", padx=10, pady=5)
text_input = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=90, height=20)
text_input.pack(padx=10, pady=5)

tk.Label(root, text="Year (e.g., 2026):").pack(anchor="w", padx=10)
year_input = tk.Entry(root)
year_input.pack(padx=10, pady=5)
year_input.insert(0, str(datetime.now().year))

tk.Button(root, text="Generate CSV", command=generate_csv, bg="green", fg="white", font=("Arial", 12)).pack(pady=10)

root.mainloop()