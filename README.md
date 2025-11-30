
import openpyxl
from openpyxl.styles import Font, Alignment

# Create a new workbook
wb = openpyxl.Workbook()

# Select the active sheet
sheet = wb.active
sheet.title = "Financial Report"

# Add headers for the columns
headers = ["Date", "Description", "Income", "Expense", "Balance"]
for col_num, header in enumerate(headers, 1):
    cell = sheet.cell(row=1, column=col_num)
    cell.value = header
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

# Function to add transaction data
def add_transaction(date, description, income, expense):
    row = sheet.max_row + 1
    sheet.cell(row=row, column=1).value = date
    sheet.cell(row=row, column=2).value = description
    sheet.cell(row=row, column=3).value = income
    sheet.cell(row=row, column=4).value = expense
    # Calculate balance automatically
    if row == 2:
        sheet.cell(row=row, column=5).value = income - expense
    else:
        sheet.cell(row=row, colu
