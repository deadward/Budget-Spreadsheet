import openpyxl
from utils import adjust_column_widths
from budget_tables import create_budget_tables
from budget_calendar import create_calendar

#comment

# Create workbook and get active sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Budget Tracker"

# Call functions to create budget tables and calendar
sheet = create_budget_tables(sheet)
sheet = create_calendar(sheet, calendar_start_row=22)  # Adjust start_row as needed

# Adjust column widths for better presentation
adjust_column_widths(sheet)

# Save the workbook
wb.save("Budget_Tracker.xlsx")
print("Budget Tracker Excel file generated as 'Budget_Tracker.xlsx'.")
