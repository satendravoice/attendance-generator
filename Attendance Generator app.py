import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# --- Streamlit UI ---
st.title("📊 Online Zoom Attendance Generator")

st.markdown("""
### 🔹 How It Works:
1. **Upload Zoom Log CSV Files** (Multiple files supported).
2. **Enter Session Details** (Start Time, End Time, and Required Attendance).
3. **Generate Attendance** ✅
4. **Download Processed Excel Files** 📂
""")

# --- File Upload Section ---
uploaded_files = st.file_uploader("Upload CSV Files", type="csv", accept_multiple_files=True)

session_start = st.text_input("Session Start Time (YYYY-MM-DD HH:MM:SS)", "2025-02-06 09:30:00")
session_end = st.text_input("Session End Time (YYYY-MM-DD HH:MM:SS)", "2025-02-06 11:30:00")
time_required = st.number_input("Required Attendance Time (minutes)", min_value=1, value=10)

# --- Helper Functions ---
def parse_datetime(dt_str):
    """Converts a string to datetime format."""
    try:
        return datetime.strptime(dt_str.strip(), '%Y-%m-%d %H:%M:%S')
    except Exception:
        st.error(f"⚠ Invalid datetime format: {dt_str}. Expected: YYYY-MM-DD HH:MM:SS")
        return None

def process_attendance(file, session_start, session_end, time_required):
    """Processes the CSV file and generates attendance."""
    df = pd.read_csv(file, skiprows=3)  # Skipping first 3 rows as per Zoom log format
    df.columns = df.columns.str.strip()  # Clean column names

    if "Name (Original Name)" in df.columns:
        name_col = "Name (Original Name)"
    elif "Name" in df.columns:
        name_col = "Name"
    else:
        st.error("❌ CSV must contain a 'Name' or 'Name (Original Name)' column.")
        return None

    if "User Email" in df.columns:
        email_col = "User Email"
    elif "Email" in df.columns:
        email_col = "Email"
    else:
        st.error("❌ CSV must contain a 'User Email' or 'Email' column.")
        return None

    for col in ["Join Time", "Leave Time"]:
        if col not in df.columns:
            st.error(f"❌ CSV must contain '{col}' column.")
            return None

    df["Join Time"] = pd.to_datetime(df["Join Time"])
    df["Leave Time"] = pd.to_datetime(df["Leave Time"])

    # Grouping by participant to merge multiple entries
    df["Name_lower"] = df[name_col].str.lower()
    attendance = {}
    for name_lower, group in df.groupby("Name_lower"):
        original_name = group.iloc[0][name_col]
        email = group.iloc[0][email_col]
        min_join = group["Join Time"].min()
        max_leave = group["Leave Time"].max()
        total_duration = (group["Leave Time"] - group["Join Time"]).sum().seconds / 60

        # Attendance status
        status = "P" if total_duration >= time_required else "A"

        attendance[name_lower] = {
            "Name": original_name,
            "Email": email,
            "Join Time": min_join.strftime('%Y-%m-%d %H:%M:%S'),
            "Leave Time": max_leave.strftime('%Y-%m-%d %H:%M:%S'),
            "Duration": round(total_duration, 2),
            "Status": status
        }

    # Convert to DataFrame
    attendance_df = pd.DataFrame.from_dict(attendance, orient="index")
    
    return attendance_df

# --- Attendance Generation ---
if st.button("Generate Attendance"):
    if not uploaded_files:
        st.error("❌ Please upload at least one CSV file.")
    else:
        processed_files = []
        summary_info = []
        
        for file in uploaded_files:
            processed_df = process_attendance(file, session_start, session_end, time_required)
            if processed_df is None:
                continue
            
            # Save to an in-memory Excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                processed_df.to_excel(writer, sheet_name="Attendance", index=False)
            output.seek(0)

            # Save file info for download
            processed_files.append((file.name, output))
            present_count = (processed_df["Status"] == "P").sum()
            absent_count = (processed_df["Status"] == "A").sum()
            summary_info.append(f"📂 {file.name}: **{present_count} Present**, {absent_count} Absent")

        st.success("✅ Attendance Generated Successfully!")

        # Show Summary
        st.subheader("📊 Attendance Summary")
        for line in summary_info:
            st.write(line)

        # Provide Download Links
        st.subheader("⬇ Download Processed Excel Files")
        for name, excel_data in processed_files:
            st.download_button(label=f"Download {name[:-4]}_processed.xlsx",
                               data=excel_data,
                               file_name=f"{name[:-4]}_processed.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
