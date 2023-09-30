import openpyxl
from openpyxl import load_workbook
from datetime import date
import tkinter as tk
from tkinter import simpledialog, messagebox

# Function to input student name and attendance status using a dialog
def input_attendance():
    student_name = simpledialog.askstring("Attendance Tracker", "Enter student name (or 'done' to finish):")
    if student_name is None or student_name.lower() == 'done':
        return None, None
    attendance_status = simpledialog.askstring("Attendance Tracker", "Enter attendance status (Present/Absent):")
    return student_name, attendance_status
# Function to load and display previous attendance records
def load_previous_attendance():
    try:
        workbook = load_workbook("attendance.xlsx")
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            date, student_name, attendance_status = row
            attendance_listbox.insert(tk.END, f"Date: {date}, Student: {student_name}, Status: {attendance_status}")
        workbook.close()
    except FileNotFoundError:
        messagebox.showinfo("Info", "No previous attendance records found.")

# Generate a unique filename based on the current date and time
today = date.today()
file_suffix = today.strftime("%Y%m%d")
file_name = f"attendance_{file_suffix}.xlsx"

