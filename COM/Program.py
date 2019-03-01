# Create ComHelper object and set license. 
# NOTE: If you're using a Professional version you'll need to put your serial key below.
import win32com.client as COM
comHelper = COM.Dispatch("GemBox.Spreadsheet.ComHelper")
comHelper.ComSetLicense("FREE-LIMITED-KEY")

# Create new ExcelFile object and add new worksheet.
workbook = COM.Dispatch("GemBox.Spreadsheet.ExcelFile")
worksheet = workbook.Worksheets.Add("Sheet1")

# Format first row columns.
header = worksheet.Cells.GetSubrange("A1", "B1")
header.Merged = True
header.Value = "GemBox.Spreadsheet COM Example"
header.Style.HorizontalAlignment = 2
header.Style.Font.Weight = 700

# Set column A width and values.
column = comHelper.GetColumn(worksheet, 0)
column.Width = 20 * 256
column.Cells.Item(1).Value = "1 + 1 ="
column.Cells.Item(2).Value = "B2 * 2 ="
column.Cells.Item(3).Value = "B3 * 120% ="
column.Cells.Item(4).Value = "SUM(B2:B4) ="

# Set column B width and formulas.
column = comHelper.GetColumn(worksheet, 1)
column.Width = 20 * 256
column.Cells.Item(1).Formula = "=1 + 1"
column.Cells.Item(2).Formula = "=B2 * 2"
column.Cells.Item(3).Formula = "=B3 * 120%"
column.Cells.Item(4).Formula = "=SUM(B2:B4)"

# Calculate all worksheet formulas.
worksheet.Calculate

# Output formula results.
print("Cell calculation results:")
for i in range(1, 4):
    print(" B" + str(i) + " = " + str(column.Cells.Item(i).Value))

# Get output path and save workbook as XLSX file.
import os
path = os.getcwd() + "\\ComExample.xlsx"

workbook.Save(path)
print("Workbook saved as '" + path + "'")