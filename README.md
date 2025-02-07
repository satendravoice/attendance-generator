# Attendance Generator Application

## Overview

The **Attendance Generator Application** is a GUI-based tool that processes Zoom log CSV files to generate accurate attendance records. It allows users to define multiple sessions, specify session start and end times, and automatically calculates participant attendance based on their duration in the session. The application supports both single and multiple CSV file processing and outputs the final attendance report in Excel format.

<img width="599" alt="image" src="https://github.com/user-attachments/assets/45cc4b57-08ed-4873-a0c4-34b703eab013" /> <img width="599" alt="image" src="https://github.com/user-attachments/assets/7354c47b-dde5-41ab-be6c-682657d7c451" />
*(Screenshot of the application interface)*

## Features
- **User-Friendly GUI**: Intuitive and easy-to-use interface.
- **Single and Multiple CSV Processing**: Choose between a single file with multiple sessions or multiple CSV files.
- **Custom Session Timings**: Users can specify session start and end times for accurate attendance tracking.
- **Automated Attendance Calculation**: Merges overlapping intervals, calculates total duration, and determines attendance status based on required participation time.
- **Real-Time Summary Report**: Displays present and absent participant counts for each session.
- **Accurate Merging of Duplicate Entries**: Ensures participant records are combined based on their name and email.
- **CSV and Excel Output**: Generates an Excel file with processed attendance data.
- **Error Handling and Validation**: Provides meaningful error messages for incorrect input formats or missing fields.
- **Customizable Duration Threshold**: Users can define the minimum required time to be marked as present.
- **Popup Notifications**: Displays summary messages and confirmations for better user experience.

## Installation
1. **Download and Install Python (if not already installed)**:  
   - Ensure Python 3.x is installed.
   - Install required dependencies using:
     ```bash
     pip install pandas openpyxl tk
     ```

2. **Download the Application Files**:
   - Clone or download the script.

3. **Run the Application**:
   - Execute the script by running:
     ```bash
     python attendance_generator.py
     ```

## Usage
### Selecting Mode
- **Single CSV File with Multiple Sessions**: Processes a single CSV file with multiple session timestamps.
- **Multiple CSV Files**: Processes multiple Zoom log files individually.

### Steps to Generate Attendance
1. **Select CSV File(s)**: Choose the Zoom log CSV file(s) containing participant records.
2. **Add Sessions**: Define session start and end times along with the required minimum participation time.
   - **Session Start**: The time when the session begins. Must be in the format `YYYY-MM-DD HH:MM:SS`. Example: `2024-02-08 10:30:00`.
   - **Session End**: The time when the session ends. Must also follow the format `YYYY-MM-DD HH:MM:SS`. Example: `2024-02-08 12:30:00`.
   - **Time Required (min)**: The minimum number of minutes a participant must be present in a session to be marked as "Present" (`P`). If a participant attends for less than this duration, they will be marked as "Absent" (`A`).
3. **Generate Attendance**: Click the `Generate Attendance` button to process the data.
4. **Save the Output**: Choose a location to save the final attendance report in Excel format.
5. **Review Summary**: A popup will display the number of present and absent participants per session.
### Selecting Mode
- **Single CSV File with Multiple Sessions**: Processes a single CSV file with multiple session timestamps.
- **Multiple CSV Files**: Processes multiple Zoom log files individually.

### Steps to Generate Attendance
1. **Select CSV File(s)**: Choose the Zoom log CSV file(s) containing participant records.
2. **Add Sessions**: Define session start and end times along with the required minimum participation time.
3. **Generate Attendance**: Click the `Generate Attendance` button to process the data.
4. **Save the Output**: Choose a location to save the final attendance report in Excel format.
5. **Review Summary**: A popup will display the number of present and absent participants per session.

## Input File Requirements
- **Zoom Log CSV Format** (with at least these columns):
  - `Name (Original Name)` or `Name`
  - `User Email` or `Email`
  - `Join Time` (Format: YYYY-MM-DD HH:MM:SS)
  - `Leave Time` (Format: YYYY-MM-DD HH:MM:SS)
  - `Duration (Minutes)`

## Output Format
- The generated attendance report includes:
  - `Name`, `Email`, `Join Time`, `Leave Time`
  - Attendance status (`P` for present, `A` for absent) per session
  - `Total Duration` of attendance in minutes
- Output is saved as an **Excel (.xlsx) file** with two sheets:
  - `Sheet1`: Raw Zoom log data
  - `Attendance`: Processed attendance data

## Error Handling
- If a required column is missing, an error message will be displayed.
- If session start time is after the end time, the user is prompted to correct it.
- If an invalid time format is entered, the application notifies the user.
- Ensures only valid, non-overlapping session data is processed.

## Credits
Developed by **Satendra Goswami**

For any questions or to connect with me, find me on Instagram: [@satendragoswamii](https://www.instagram.com/satendragoswamii)

