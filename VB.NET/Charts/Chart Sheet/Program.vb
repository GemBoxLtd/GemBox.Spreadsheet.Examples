Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Data")

        ' Add data which is used by the Excel chart.
        worksheet.Cells("A1").Value = "Name"
        worksheet.Cells("A2").Value = "John Doe"
        worksheet.Cells("A3").Value = "Fred Nurk"
        worksheet.Cells("A4").Value = "Hans Meier"
        worksheet.Cells("A5").Value = "Ivan Horvat"

        worksheet.Cells("B1").Value = "Salary"
        worksheet.Cells("B2").Value = 3600
        worksheet.Cells("B3").Value = 2580
        worksheet.Cells("B4").Value = 3200
        worksheet.Cells("B5").Value = 4100

        ' Set header row and formatting.
        worksheet.Rows(0).Style.Font.Weight = ExcelFont.BoldWeight
        worksheet.Columns(0).SetWidth(3, LengthUnit.Centimeter)
        worksheet.Columns(1).Style.NumberFormat = """$""#,##0"

        ' Make entire sheet print on a single page.
        worksheet.PrintOptions.FitWorksheetWidthToPages = 1
        worksheet.PrintOptions.FitWorksheetHeightToPages = 1

        ' Create Excel chart sheet.
        Dim chartsheet = workbook.Worksheets.Add(SheetType.Chart, "Chart")
        workbook.Worksheets.ActiveWorksheet = chartsheet

        ' Create Excel chart and select data for it.
        ' You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        Dim chart = chartsheet.Charts.Add(ChartType.Pie, 0, 0, 0, 0, LengthUnit.Centimeter)
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, 4, 1), True)

        workbook.Save("Chart Sheet.xlsx")
    End Sub
End Module
