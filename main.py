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

