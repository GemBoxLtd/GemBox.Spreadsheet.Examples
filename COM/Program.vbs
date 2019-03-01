' Create ComHelper object and set license. 
' NOTE: If you're using a Professional version you'll need to put your serial key below.
Set comHelper = CreateObject("GemBox.Spreadsheet.ComHelper")
comHelper.ComSetLicense("FREE-LIMITED-KEY")

' Create new ExcelFile object and add new worksheet.
Set workbook = CreateObject("GemBox.Spreadsheet.ExcelFile")
Set worksheet = workbook.Worksheets.Add("Sheet1")

' Format first row columns.
Set header = worksheet.Cells.GetSubrange("A1", "B1")
header.Merged = true
header.Value = "GemBox.Spreadsheet COM Example"
header.Style.HorizontalAlignment = 2
header.Style.Font.Weight = 700

' Set column A width and values.
Set column = comHelper.GetColumn(worksheet, 0)
column.Width = 20 * 256
column.Cells.Item(1).Value = "1 + 1 ="
column.Cells.Item(2).Value = "B2 * 2 ="
column.Cells.Item(3).Value = "B3 * 120% ="
column.Cells.Item(4).Value = "SUM(B2:B4) ="

' Set column B width and formulas.
Set column = comHelper.GetColumn(worksheet, 1)
column.Width = 20 * 256
column.Cells.Item(1).Formula = "=1 + 1"
column.Cells.Item(2).Formula = "=B2 * 2"
column.Cells.Item(3).Formula = "=B3 * 120%"
column.Cells.Item(4).Formula = "=SUM(B2:B4)"

' Calculate all worksheet formulas.
worksheet.Calculate()

' Output formula results.
Response.Write("Cell calculation results:")
For i = 1 to 4
    Response.Write(" B" & i & " = " & column.Cells.Item(i).Value)
Next

' Get output path and save workbook as XLSX file.
Dim path
path = Server.MapPath(".") & "\ComExample.xlsx"

workbook.Save(path)
Response.Write("Workbook saved as '" & path & "'")