import os
import win32com.client as COM

# Create ComHelper object.
comHelper = COM.Dispatch("GemBox.Spreadsheet.ComHelper")
# If using the Professional version, put your serial key below.
comHelper.ComSetLicense("FREE-LIMITED-KEY")

####################
### Create Excel ###
####################

# Create new ExcelFile object.
workbook = COM.Dispatch("GemBox.Spreadsheet.ExcelFile")
# Add new ExcelWorksheet object.
worksheet = workbook.Worksheets.Add("Sheet1")

# Set width and format of column "A".
columnA = comHelper.GetColumn(worksheet, 0)
columnA.Width = 20 * 256
columnA.Style.Font.Weight = 700

# Set values of cells "A1", "A2", "A3" and "A4".
columnA.Cells.Item(0).Value = "John Doe"
columnA.Cells.Item(1).Value = "Bob Garvey"
columnA.Cells.Item(2).Value = "Ben Stilwell"
columnA.Cells.Item(3).Value = "Peter Pan"
  
# Set values of cells "B1", "B2", "B3" and "B4".
columnB = comHelper.GetColumn(worksheet, 1)
columnB.Cells.Item(0).Value = 1000
columnB.Cells.Item(1).Value = 2000
columnB.Cells.Item(2).Value = 3000
columnB.Cells.Item(3).Value = 4000

# Create new Excel file.
workbook.Save(os.getcwd() + "\\New.xlsx")

##################
### Read Excel ###
##################

# Read existing Excel file.
book = comHelper.Load(os.getcwd() + "\\New.xlsx")
# Get first Excel sheet.
sheet = book.Worksheets.Item(0)
# Get first Excel row.
row1 = comHelper.GetRow(sheet, 0)

# Display values of cells "A1" and "B1".
print("Cell A1:" + str(row1.Cells.Item(0).Value))
print("<br>")
print("Cell B1:" + str(row1.Cells.Item(1).Value))

####################
### Update Excel ###
####################

# Update values of cells "A1" and "B1".
row1.Cells.Item(0).Value = "Jane Doe"
row1.Cells.Item(1).Value = 2000

# Write the updated Excel file.
book.Save(os.getcwd() + "\\Updated.xlsx")