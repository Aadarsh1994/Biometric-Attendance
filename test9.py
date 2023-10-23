import openpyxl
import random
from datetime import datetime, timedelta
import pandas as pd

# Load student information from an external Excel sheet
student_info_file = "testt.xlsx"  # Replace with your file path
student_info = pd.read_excel(student_info_file)

# Define the date range for two months
start_date = datetime(2015, 12, 21)  ##yyyy, m, d
end_date = datetime(2016, 2, 20)    ##yyyy, m, d

# Define the time range for Time IN and Time OUT
time_in_start =  datetime(2015, 12, 21, 8, 00)     #start_date
time_in_end =    datetime(2015, 12, 21, 8, 14)        #start_date
time_out_start = datetime(2015, 12, 21, 18, 11)   #start_date
time_out_end =   datetime(2015, 12, 21, 18, 31)     #start_date

# Create a new Excel workbook to store attendance records
workbook = openpyxl.Workbook()
del workbook["Sheet"]  # Remove the default sheet

for date in pd.date_range(start=start_date, end=end_date):
    if date.weekday() == 6:  # Skip Sundays (Sunday is represented as 6)
        continue

    # Create a new worksheet for each day's attendance
    attendance_worksheet = workbook.create_sheet(title=date.strftime('%Y-%m-%d'))

    # Add headers to the worksheet, including "Status"
    headers = ["Sl.No", "STUDENT ID", "STUDENT NAME", "ATTENDANCE DATE", "ATTENDANCE DAY", "IN TIME", "OUT TIME", "STATUS"]
    attendance_worksheet.append(headers)

    # Initialize the Serial Number for each sheet
    serial_number = 1

    # Generate random attendance records for all students
    for _, student in student_info.iterrows():
        student_id = student["Student ID"]
        student_name = student["Student Name"]

        # Generate random time within the specified ranges
        time_in = time_in_start + timedelta(seconds=random.randint(0, 840))  # Random time within 30 minutes
        time_out = time_out_start + timedelta(seconds=random.randint(0, 1200))  # Random time within 1 hour

        # Format the date and day
        attendance_date = date.strftime('%Y-%m-%d')
        attendance_day = date.strftime('%A')

        # Format time_in and time_out as strings without the date
        time_in_str = time_in.strftime('%H:%M:%S')
        time_out_str = time_out.strftime('%H:%M:%S')

        # Set the status as "Present"
        status = "Present"

        # Create an attendance row with the Serial Number
        attendance_row = [serial_number, student_id, student_name, attendance_date, attendance_day, time_in_str, time_out_str, status]
        attendance_worksheet.append(attendance_row)

        # Increment the Serial Number for the next student
        serial_number += 1

# Save the entire Excel workbook
workbook.save("Batch_4001.xlsx")
