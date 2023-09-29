# to save a excel file/sheet for this 
import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Create a new sheet
sheet = workbook.active

# Add headers
sheet["A1"] = "Date"
sheet["B1"] = "Student Name"
sheet["C1"] = "Attendance Status"

# Save the workbook
workbook.save("attendance.xlsx")

