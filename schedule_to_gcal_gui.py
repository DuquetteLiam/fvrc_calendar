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
    """Return start_time, end_time in HH:MM 24h format; end_time can be None.
    Searches anywhere in the string so that times embedded in descriptions are
    detected. If no valid range is found, returns (None, None)."""
    text = text.replace("–", "-")
    time_pattern = r'(\d{1,2}(:\d{2})?)\s*(am|pm)?\s*(?:-\s*(\d{1,2}(:\d{2})?)\s*(am|pm)?)?'
    match = re.search(time_pattern, text.strip(), re.IGNORECASE)
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
        return h, m, f"{h:02d}:{m:02d}"
    
    start_h, start_m, start24 = to_24h(start, start_ampm)
    
    # if end has am/pm and start doesn't, match start to end
    if end_ampm and not start_ampm:
        if end_ampm.lower() == 'am':
            # if start > end hour, wrap into same morning
            if start_h > int(end.split(':')[0]):
                # keep as-is (e.g. 3-5am) start_h already <12
                pass
            start24 = f"{start_h:02d}:{start_m:02d}"
        else:
            # pm case
            if start_h < 12:
                start_h += 12
                start24 = f"{start_h:02d}:{start_m:02d}"
    # if still no am/pm, assume pm for early afternoon/evening times 1-8
    elif not start_ampm and 1 <= start_h <= 8:
        start_h += 12
        start24 = f"{start_h:02d}:{start_m:02d}"

    if end:
        end_h, end_m, end24 = to_24h(end, end_ampm)
        # if end is not explicitly marked and is less than start, assume it's PM
        if not end_ampm and end_h < 12 and end_h < start_h:
            end_h += 12
            end24 = f"{end_h:02d}:{end_m:02d}"
    else:
        end24 = None
    
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

    # common time range regex used for splitting and stripping
    time_range_pattern = r"\d{1,2}(?::\d{2})?\s*(?:am|pm)?\s*-\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)?"

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

        # Detect day lines (e.g. '23 Mon', '24 Tues', '5 Friday') and capture any trailing events
        # only match if the word after the number is a weekday name/abbreviation
        day_match = re.match(
            r'(\d{1,2})\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\w*(.*)',
            line, re.IGNORECASE)
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

        # helper that splits line into event segments, even when adjacent times are smushed together
        def split_events(text):
            # first break on bullets or two-or-more spaces
            parts = re.split(r'[•·]| {2,}', text)
            out = []
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                matches = list(re.finditer(time_range_pattern, part, re.IGNORECASE))
                if len(matches) <= 1:
                    out.append(part)
                else:
                    starts = [m.start() for m in matches]
                    for idx, start in enumerate(starts):
                        end = starts[idx+1] if idx+1 < len(starts) else len(part)
                        seg = part[start:end].strip(' ·•')
                        if seg:
                            out.append(seg)
            return out

        # Split multiple events in one line (handles smushed entries too)
        items = split_events(line)
        for item in items:
            item = item.strip()
            if not item:
                continue
            start_time, end_time = parse_time_range(item)
            # decide how to treat the description and all-day flag
            if start_time and not end_time:
                # only a start is listed; keep the time in the title and mark as all-day
                desc = item
                all_day = "True"
            elif start_time:
                # full range known, strip it from the text
                desc = re.sub(r'^' + time_range_pattern + r'\s*', '', item, flags=re.IGNORECASE).strip()
                all_day = "False"
            else:
                desc = item
                all_day = "True"
            if not desc:
                desc = item
            # For start-only (all_day==True and original had a start), do NOT include start/end times in CSV
            csv_start_time = "" if (all_day == "True" and start_time) else (start_time if start_time else "")
            csv_end_date = current_date.strftime("%m/%d/%Y") if (end_time and all_day == "False") else ""
            csv_end_time = end_time if (end_time and all_day == "False") else ""
            events.append({
                "Subject": desc,
                "Start Date": current_date.strftime("%m/%d/%Y"),
                "Start Time": csv_start_time,
                "End Date": csv_end_date,
                "End Time": csv_end_time,
                "All Day Event": all_day,
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
    result = {"ok": False, "events": events}
    preview = tk.Toplevel()
    preview.title("Preview Events — Edit as needed")
    # Use a larger default size and a reasonable minimum so buttons are visible
    preview.geometry("800x500")
    preview.minsize(640, 380)

    tk.Label(preview, text="Preview — edit and confirm to generate CSV", font=(None, 11, "bold")).pack(anchor="w", padx=8, pady=(6,0))
    text = scrolledtext.ScrolledText(preview, wrap=tk.WORD)
    text.pack(expand=True, fill=tk.BOTH, padx=8, pady=(4,0))
    for e in events:
        start_time = e["Start Time"] if e["Start Time"] else ""
        end_time = e["End Time"] if e["End Time"] else ""
        line = f'{e["Start Date"]} {start_time}-{end_time} {e["Subject"]}\n'
        text.insert(tk.END, line)

    def parse_preview_text():
        """Parse the edited text back into events"""
        text_content = text.get("1.0", tk.END)
        lines = text_content.strip().split('\n')
        parsed_events = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Parse format: MM/DD/YYYY HH:MM-HH:MM Title
            # or: MM/DD/YYYY Title (no times)
            match = re.match(r'(\d{1,2}/\d{1,2}/\d{4})\s+((\d{1,2}:\d{2})?-?(\d{1,2}:\d{2})?)?\s+(.*)', line)
            if match:
                date_str = match.group(1)
                time_range = match.group(2) if match.group(2) else ""
                subject = match.group(5).strip()
                
                start_time = ""
                end_time = ""
                all_day = "True"
                
                if time_range and "-" in time_range:
                    parts = time_range.split('-')
                    start_time = parts[0].strip()
                    end_time = parts[1].strip() if len(parts) > 1 else ""
                    all_day = "False"
                elif time_range:
                    start_time = time_range.strip()
                    all_day = "False"
                
                parsed_events.append({
                    "Subject": subject,
                    "Start Date": date_str,
                    "Start Time": start_time,
                    "End Date": date_str if end_time else "",
                    "End Time": end_time,
                    "All Day Event": all_day,
                    "Description": "",
                    "Location": ""
                })
        
        return parsed_events if parsed_events else None

    def on_generate():
        parsed = parse_preview_text()
        if parsed is None:
            messagebox.showerror("Error", "Could not parse edited events. Check the format.")
            return
        result["ok"] = True
        result["events"] = parsed
        preview.destroy()

    def on_back():
        preview.destroy()

    btn_frame = tk.Frame(preview)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=8)
    tk.Button(btn_frame, text="Generate CSV", width=15, bg="green", fg="white", font=(None, 10, "bold"), command=on_generate).pack(side=tk.LEFT, padx=8)
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
    return result["ok"], result["events"]

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
    ok, events = preview_events(events)
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