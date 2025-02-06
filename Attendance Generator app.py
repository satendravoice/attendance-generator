pip install pandas openpyxl
```)  

Below is the complete updated code:

---

```python
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

# ---------------------------
# Helper Functions
# ---------------------------
def parse_datetime(dt_str):
    """Parse a datetime string in the format: YYYY-MM-DD HH:MM:SS"""
    try:
        return datetime.strptime(dt_str.strip(), '%Y-%m-%d %H:%M:%S')
    except Exception:
        raise ValueError(f"Invalid datetime format: {dt_str}. Expected format: YYYY-MM-DD HH:MM:SS")

def merge_intervals(intervals):
    """Merge overlapping intervals (list of (start, end) datetime tuples)."""
    if not intervals:
        return []
    intervals.sort(key=lambda x: x[0])
    merged = [intervals[0]]
    for current in intervals[1:]:
        last = merged[-1]
        if current[0] <= last[1]:
            merged[-1] = (last[0], max(last[1], current[1]))
        else:
            merged.append(current)
    return merged

def compute_total_duration(intervals):
    """Return total duration in minutes from a list of (start, end) intervals."""
    total = timedelta()
    for start, end in intervals:
        total += (end - start)
    return total.total_seconds() / 60

def intersect_interval(interval, period):
    """Return the overlapping portion of an interval with a given period."""
    start, end = interval
    p_start, p_end = period
    new_start = max(start, p_start)
    new_end = min(end, p_end)
    if new_start < new_end:
        return (new_start, new_end)
    return None

def get_global_times(file_path):
    """
    Read the CSV file (skip the first three rows) and compute for each participant
    the overall (raw) minimum Join Time and maximum Leave Time.
    Returns a dict mapping lower-case name to (global_join, global_leave).
    """
    try:
        df = pd.read_csv(file_path, skiprows=3)
        df.columns = df.columns.str.strip()
    except Exception as e:
        raise ValueError(f"Error reading file '{file_path}' for global times: {e}")

    if "Name" in df.columns:
        name_col = "Name"
    elif "Name (Original Name)" in df.columns:
        name_col = "Name (Original Name)"
    else:
        raise ValueError(f"CSV file '{file_path}' must contain a 'Name' or 'Name (Original Name)' column for global times.")
    try:
        df["Join Time"] = pd.to_datetime(df["Join Time"])
        df["Leave Time"] = pd.to_datetime(df["Leave Time"])
    except Exception as e:
        raise ValueError(f"Error converting join/leave times in '{file_path}' for global times: {e}")
    df["Name_lower"] = df[name_col].str.lower()
    global_times = {}
    for name_lower, group in df.groupby("Name_lower"):
        global_join = group["Join Time"].min()
        global_leave = group["Leave Time"].max()
        global_times[name_lower] = (global_join, global_leave)
    return global_times

def process_csv_session(file_path, session_start, session_end, time_required):
    """
    Process one CSV file for a single session.
    Reads CSV with skiprows=3 (header on row 4) and expects:
      - Name: "Name (Original Name)" or "Name"
      - Email: "User Email" or "Email"
      - "Join Time" and "Leave Time"
    For each participant:
      - Merges overlapping intervals.
      - Considers only portions overlapping the session period.
      - Computes session duration (from overlapping portions).
      - Determines status ("P" if duration >= time_required, else "A").
    Returns a dict mapping lower-case name to:
      {
         "Name": original name,
         "Email": email,
         "session_min_join": earliest join from overlapping portion,
         "raw_max_leave": maximum leave from overlapping portion,
         "session_duration": total minutes attended in session,
         "status": "P" or "A"
      }
    """
    try:
        df = pd.read_csv(file_path, skiprows=3)
        df.columns = df.columns.str.strip()
    except Exception as e:
        raise ValueError(f"Error reading file '{file_path}': {e}")
    if "Name" in df.columns:
        name_col = "Name"
    elif "Name (Original Name)" in df.columns:
        name_col = "Name (Original Name)"
    else:
        raise ValueError(f"CSV file '{file_path}' must contain a 'Name' or 'Name (Original Name)' column.")
    if "Email" in df.columns:
        email_col = "Email"
    elif "User Email" in df.columns:
        email_col = "User Email"
    else:
        raise ValueError(f"CSV file '{file_path}' must contain an 'Email' or 'User Email' column.")
    for col in ["Join Time", "Leave Time"]:
        if col not in df.columns:
            raise ValueError(f"CSV file '{file_path}' must contain a '{col}' column.")
    try:
        df["Join Time"] = pd.to_datetime(df["Join Time"])
        df["Leave Time"] = pd.to_datetime(df["Leave Time"])
    except Exception as e:
        raise ValueError(f"Error converting join/leave times in '{file_path}': {e}")
    df["Name_lower"] = df[name_col].str.lower()
    session_results = {}
    for name_lower, group in df.groupby("Name_lower"):
        original_name = group.iloc[0][name_col]
        email = group.iloc[0][email_col]
        intervals = [(row["Join Time"], row["Leave Time"]) for _, row in group.iterrows()]
        merged = merge_intervals(intervals)
        session_intervals = [intersect_interval(interval, (session_start, session_end))
                             for interval in merged if intersect_interval(interval, (session_start, session_end)) is not None]
        session_intervals = merge_intervals(session_intervals)
        session_duration = compute_total_duration(session_intervals)
        if session_intervals:
            session_min_join = min(i[0] for i in session_intervals)
            raw_max_leave = max(i[1] for i in session_intervals)
            status = "P" if session_duration >= time_required else "A"
        else:
            continue
        session_results[name_lower] = {
            "Name": original_name,
            "Email": email,
            "session_min_join": session_min_join,
            "raw_max_leave": raw_max_leave,
            "session_duration": session_duration,
            "status": status
        }
    return session_results

# ---------------------------
# Main Application Class with Compact & Advanced UI
# ---------------------------
class AttendanceApp:
    def __init__(self, master):
        self.master = master
        master.title("Advanced Zoom Attendance Generator")
        master.geometry("800x600")
        master.resizable(False, False)
        self.mode = tk.StringVar(value="single")  # "single" or "multiple"

        style = ttk.Style()
        style.theme_use("clam")

        # Main container frame
        self.main_frame = ttk.Frame(master, padding="5")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        # Row 0: Mode Selection
        mode_frame = ttk.LabelFrame(self.main_frame, text="Select Mode", padding="5")
        mode_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="Single CSV File with Multiple Sessions",
                        variable=self.mode, value="single", command=self.switch_mode).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Radiobutton(mode_frame, text="Multiple CSV Files (Generate Excel per File)",
                        variable=self.mode, value="multiple", command=self.switch_mode).grid(row=0, column=1, padx=5, pady=2, sticky="w")

        # Row 1: File Selection
        file_frame = ttk.LabelFrame(self.main_frame, text="CSV File Selection", padding="5")
        file_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        self.file_select_button = ttk.Button(file_frame, text="Select CSV File", command=self.select_files)
        self.file_select_button.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.selected_files_label = ttk.Label(file_frame, text="No file selected")
        self.selected_files_label.grid(row=0, column=1, padx=5, pady=2)

        # For multiple mode, maintain a mapping of file names to paths.
        self.file_mapping = {}
        self.selected_file = None
        self.selected_files = []

        # Row 2: Sessions Section (for per‑session settings)
        self.sessions_outer = ttk.LabelFrame(self.main_frame, text="Session Details", padding="5")
        self.sessions_outer.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.header_frame = ttk.Frame(self.sessions_outer)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=3, pady=2)
        ttk.Label(self.header_frame, text="File", width=20, anchor="w", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, padx=3)
        ttk.Label(self.header_frame, text="Session Start", width=20, anchor="center", font=('Segoe UI', 9, 'bold')).grid(row=0, column=1, padx=3)
        ttk.Label(self.header_frame, text="Session End", width=20, anchor="center", font=('Segoe UI', 9, 'bold')).grid(row=0, column=2, padx=3)
        ttk.Label(self.header_frame, text="Time Required (min)", width=18, anchor="center", font=('Segoe UI', 9, 'bold')).grid(row=0, column=3, padx=3)
        self.sessions_container = ttk.Frame(self.sessions_outer)
        self.sessions_container.grid(row=1, column=0, sticky="ew", padx=3, pady=2)
        self.session_rows = []

        # Row 3: Add Session Button (enabled in both modes)
        self.add_session_button = ttk.Button(self.main_frame, text="Add Session", command=self.add_session)
        self.add_session_button.grid(row=3, column=0, sticky="w", padx=10, pady=2)

        # Row 4: Generate Button
        generate_frame = ttk.Frame(self.main_frame, padding="5")
        generate_frame.grid(row=4, column=0, sticky="e", padx=5, pady=2)
        self.generate_button = ttk.Button(generate_frame, text="Generate Attendance", command=self.generate_attendance)
        self.generate_button.grid(row=0, column=0, padx=5)

        # Row 5: Watermark
        self.watermark = ttk.Label(self.main_frame, text="App created by Satendra Goswami", foreground="gray", font=('Segoe UI', 8))
        self.watermark.grid(row=5, column=0, pady=3)

        self.switch_mode()

    def switch_mode(self):
        mode = self.mode.get()
        if mode == "single":
            self.file_select_button.config(text="Select CSV File")
            self.selected_files_label.config(text="(Only one file allowed)")
            self.selected_file = None
            self.selected_files = []
            for child in self.sessions_container.winfo_children():
                child.destroy()
            self.session_rows = []
        else:
            self.file_select_button.config(text="Select CSV Files")
            self.selected_files_label.config(text="No files selected")
            self.selected_files = []
            self.selected_file = None
            for child in self.sessions_container.winfo_children():
                child.destroy()
            self.session_rows = []

    def select_files(self):
        mode = self.mode.get()
        if mode == "single":
            file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select Zoom Log CSV File")
            if file_path:
                self.selected_file = file_path
                self.selected_files_label.config(text=os.path.basename(file_path))
                for child in self.sessions_container.winfo_children():
                    child.destroy()
                self.session_rows = []
        else:
            files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")], title="Select Zoom Log CSV Files")
            if files:
                self.selected_files = list(files)
                self.file_mapping = {os.path.basename(f): f for f in self.selected_files}
                self.selected_files_label.config(text=f"{len(self.selected_files)} file(s) selected")
                for child in self.sessions_container.winfo_children():
                    child.destroy()
                self.session_rows = []

    def add_session(self):
        row_frame = ttk.Frame(self.sessions_container)
        row_frame.grid(sticky="ew", padx=3, pady=2)
        if self.mode.get() == "single":
            file_name = os.path.basename(self.selected_file) if self.selected_file else ""
            file_widget = ttk.Label(row_frame, text=file_name, width=20, anchor="w")
            file_widget.grid(row=0, column=0, padx=3)
            session_dict = {"file_path": self.selected_file}
        else:
            file_options = list(self.file_mapping.keys())
            file_var = tk.StringVar(value=file_options[0] if file_options else "")
            file_widget = ttk.Combobox(row_frame, textvariable=file_var, values=file_options, state="readonly", width=20)
            file_widget.grid(row=0, column=0, padx=3)
            session_dict = {"file_var": file_var}
        start_entry = ttk.Entry(row_frame, width=20)
        start_entry.grid(row=0, column=1, padx=3)
        end_entry = ttk.Entry(row_frame, width=20)
        end_entry.grid(row=0, column=2, padx=3)
        time_req_entry = ttk.Entry(row_frame, width=18)
        time_req_entry.grid(row=0, column=3, padx=3)
        session_dict.update({
            "start_entry": start_entry,
            "end_entry": end_entry,
            "time_required_entry": time_req_entry
        })
        self.session_rows.append(session_dict)

    def generate_attendance(self):
        mode = self.mode.get()
        # In single mode, process as before.
        if mode == "single":
            if not self.selected_file:
                messagebox.showerror("Error", "Please select a CSV file in Single CSV mode.")
                return
            if not self.session_rows:
                messagebox.showerror("Error", "Please add at least one session.")
                return
            sessions_info = []
            for sess in self.session_rows:
                s_text = sess["start_entry"].get().strip()
                e_text = sess["end_entry"].get().strip()
                tr_text = sess["time_required_entry"].get().strip()
                if not s_text or not e_text or not tr_text:
                    messagebox.showerror("Error", "Session details missing in one of the session rows.")
                    return
                try:
                    s_dt = parse_datetime(s_text)
                    e_dt = parse_datetime(e_text)
                    time_req_val = float(tr_text)
                except ValueError as ve:
                    messagebox.showerror("Error", f"Error in session details: {ve}")
                    return
                if s_dt >= e_dt:
                    messagebox.showerror("Error", "Session Start must be before Session End.")
                    return
                sessions_info.append({
                    "file_path": sess["file_path"],
                    "session_start": s_dt,
                    "session_end": e_dt,
                    "time_required": time_req_val
                })
            # Process single file as before.
            global_participants = {}
            total_sessions = len(sessions_info)
            session_labels = []
            for session_index, session in enumerate(sessions_info, start=1):
                try:
                    session_results = process_csv_session(session["file_path"], session["session_start"], session["session_end"], session["time_required"])
                    session_global_times = get_global_times(session["file_path"])
                except ValueError as ve:
                    messagebox.showerror("Error", str(ve))
                    return
                for name_lower, details in session_results.items():
                    if name_lower in session_global_times:
                        g_join, g_leave = session_global_times[name_lower]
                    else:
                        continue
                    if name_lower not in global_participants:
                        global_participants[name_lower] = {
                            "Name": details["Name"],
                            "Email": details["Email"],
                            "global_join": g_join,
                            "global_leave": g_leave,
                            "total_duration": details["session_duration"],
                            "sessions": {session_index: details["status"]}
                        }
                    else:
                        curr = global_participants[name_lower]
                        curr["global_join"] = min(curr["global_join"], g_join)
                        curr["global_leave"] = max(curr["global_leave"], g_leave)
                        curr["total_duration"] += details["session_duration"]
                        curr["sessions"][session_index] = details["status"]
                session_labels.append(f"Session {session_index} ({session['session_start'].strftime('%Y-%m-%d %H:%M:%S')})")
            for participant in global_participants.values():
                for i in range(1, total_sessions + 1):
                    if i not in participant["sessions"]:
                        participant["sessions"][i] = "A"
            output_records = []
            for participant in global_participants.values():
                record = {
                    "Name": participant["Name"],
                    "Email": participant["Email"],
                    "Join Time": participant["global_join"].strftime('%Y-%m-%d %H:%M:%S'),
                    "Leave Time": participant["global_leave"].strftime('%Y-%m-%d %H:%M:%S')
                }
                for i in range(1, total_sessions + 1):
                    record[session_labels[i-1]] = participant["sessions"][i]
                record["Duration"] = round(participant["total_duration"], 2)
                output_records.append(record)
            attendance_df = pd.DataFrame(output_records)
            try:
                raw_log_df = pd.read_csv(self.selected_file)
            except Exception as e:
                messagebox.showerror("Error", f"Error reading raw log from '{self.selected_file}': {e}")
                return
            base = os.path.basename(self.selected_file)
            name_part, _ = os.path.splitext(base)
            output_file = filedialog.asksaveasfilename(initialfile=name_part + "_processed.xlsx",
                                                       defaultextension=".xlsx",
                                                       filetypes=[("Excel files", "*.xlsx")],
                                                       title="Save Attendance Excel File as")
            if not output_file:
                return
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    raw_log_df.to_excel(writer, sheet_name="Sheet1", index=False)
                    attendance_df.to_excel(writer, sheet_name="Attendance", index=False)
            except Exception as e:
                messagebox.showerror("Error", f"Error saving output Excel file: {e}")
                return
            # Build summary for single file.
            session_summary = []
            for i in range(1, total_sessions + 1):
                present_count = sum(1 for p in global_participants.values() if p["sessions"].get(i, "A") == "P")
                absent_count = sum(1 for p in global_participants.values() if p["sessions"].get(i, "A") == "A")
                session_summary.append(f"Session {i}: Present: {present_count}, Absent: {absent_count}")
            summary_str = "\n".join(session_summary)
            messagebox.showinfo("Attendance Generated",
                                f"Attendance generated and saved to:\n{output_file}\n\nAttendance Summary:\n{summary_str}")
        else:
            # Multiple mode: Process each file individually.
            # Build a dictionary: file_path -> list of sessions for that file.
            sessions_by_file = {}
            for sess in self.session_rows:
                s_text = sess["start_entry"].get().strip()
                e_text = sess["end_entry"].get().strip()
                tr_text = sess["time_required_entry"].get().strip()
                if not s_text or not e_text or not tr_text:
                    messagebox.showerror("Error", "Session details missing in one of the session rows.")
                    return
                try:
                    s_dt = parse_datetime(s_text)
                    e_dt = parse_datetime(e_text)
                    time_req_val = float(tr_text)
                except ValueError as ve:
                    messagebox.showerror("Error", f"Error in session details: {ve}")
                    return
                if s_dt >= e_dt:
                    messagebox.showerror("Error", "Session Start must be before Session End.")
                    return
                file_name = sess["file_var"].get()
                if file_name not in self.file_mapping:
                    messagebox.showerror("Error", f"Selected file '{file_name}' not found.")
                    return
                file_path = self.file_mapping[file_name]
                if file_path not in sessions_by_file:
                    sessions_by_file[file_path] = []
                sessions_by_file[file_path].append({
                    "session_start": s_dt,
                    "session_end": e_dt,
                    "time_required": time_req_val
                })
            if not sessions_by_file:
                messagebox.showerror("Error", "No session information available.")
                return
            summary_all = {}
            # Process each file individually.
            for file_path, sessions_info in sessions_by_file.items():
                global_participants = {}
                total_sessions = len(sessions_info)
                session_labels = []
                for session_index, session in enumerate(sessions_info, start=1):
                    try:
                        session_results = process_csv_session(file_path, session["session_start"], session["session_end"], session["time_required"])
                        session_global_times = get_global_times(file_path)
                    except ValueError as ve:
                        messagebox.showerror("Error", str(ve))
                        return
                    for name_lower, details in session_results.items():
                        if name_lower in session_global_times:
                            g_join, g_leave = session_global_times[name_lower]
                        else:
                            continue
                        if name_lower not in global_participants:
                            global_participants[name_lower] = {
                                "Name": details["Name"],
                                "Email": details["Email"],
                                "global_join": g_join,
                                "global_leave": g_leave,
                                "total_duration": details["session_duration"],
                                "sessions": {session_index: details["status"]}
                            }
                        else:
                            curr = global_participants[name_lower]
                            curr["global_join"] = min(curr["global_join"], g_join)
                            curr["global_leave"] = max(curr["global_leave"], g_leave)
                            curr["total_duration"] += details["session_duration"]
                            curr["sessions"][session_index] = details["status"]
                    session_labels.append(f"Session {session_index} ({session['session_start'].strftime('%Y-%m-%d %H:%M:%S')})")
                for participant in global_participants.values():
                    for i in range(1, total_sessions + 1):
                        if i not in participant["sessions"]:
                            participant["sessions"][i] = "A"
                output_records = []
                for participant in global_participants.values():
                    record = {
                        "Name": participant["Name"],
                        "Email": participant["Email"],
                        "Join Time": participant["global_join"].strftime('%Y-%m-%d %H:%M:%S'),
                        "Leave Time": participant["global_leave"].strftime('%Y-%m-%d %H:%M:%S')
                    }
                    for i in range(1, total_sessions + 1):
                        record[session_labels[i-1]] = participant["sessions"][i]
                    record["Duration"] = round(participant["total_duration"], 2)
                    output_records.append(record)
                attendance_df = pd.DataFrame(output_records)
                try:
                    raw_log_df = pd.read_csv(file_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading raw log from '{file_path}': {e}")
                    return
                base = os.path.basename(file_path)
                name_part, _ = os.path.splitext(base)
                output_file = os.path.join(os.path.dirname(file_path), name_part + "_processed.xlsx")
                try:
                    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                        raw_log_df.to_excel(writer, sheet_name="Sheet1", index=False)
                        attendance_df.to_excel(writer, sheet_name="Attendance", index=False)
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving output Excel file for {base}: {e}")
                    return
                # Build session summary for this file.
                session_summary = []
                for i in range(1, total_sessions + 1):
                    present_count = sum(1 for p in global_participants.values() if p["sessions"].get(i, "A") == "P")
                    absent_count = sum(1 for p in global_participants.values() if p["sessions"].get(i, "A") == "A")
                    session_summary.append(f"Session {i}: Present: {present_count}, Absent: {absent_count}")
                summary_all[base] = "\n".join(session_summary)
            summary_str = "\n\n".join([f"{fname}:\n{summary}" for fname, summary in summary_all.items()])
            messagebox.showinfo("Attendance Generated",
                                f"Processed {len(summary_all)} files.\n\nSummary:\n{summary_str}")

# ---------------------------
# Main Program Entry
# ---------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
