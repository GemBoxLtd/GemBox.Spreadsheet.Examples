Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile

        Dim numberOfEmployees As Integer = 4

        Dim worksheet = workbook.Worksheets.Add("SourceSheet")

        ' Add data which is used by the Excel chart.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1

            worksheet.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))
            worksheet.Cells(i + 1, 1).SetValue(random.Next(1000, 5000))
        Next

        ' Set header row and formatting.
        worksheet.Cells(0, 0).Value = "Name"
        worksheet.Cells(0, 1).Value = "Salary"
        worksheet.Cells(0, 0).Style.Font.Weight = ExcelFont.BoldWeight
        worksheet.Cells(0, 1).Style.Font.Weight = ExcelFont.BoldWeight
        worksheet.Columns(0).Width = CInt(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart))
        worksheet.Columns(1).Style.NumberFormat = """$""#,##0"

        ' Create Excel chart sheet.
        Dim chartsheet = workbook.Worksheets.Add(SheetType.Chart, "ChartSheet")

        ' Create Excel chart and select data for it.
        ' You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        Dim chart = chartsheet.Charts.Add(ChartType.Bar, 0, 0, 10, 10, LengthUnit.Centimeter)
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), True)

        workbook.Save("Chart Sheet.xlsx")
    End Sub
End Module