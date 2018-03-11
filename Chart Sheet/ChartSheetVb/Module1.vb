Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile

        Dim numberOfEmployees As Integer = 4

        Dim ws1 = ef.Worksheets.Add("SourceSheet")

        ' Add data which is used by the Excel chart.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1
            ws1.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))
            ws1.Cells(i + 1, 1).SetValue(random.Next(1000, 5000))
        Next

        ' Set header row and formatting.
        ws1.Cells(0, 0).Value = "Name"
        ws1.Cells(0, 1).Value = "Salary"
        ws1.Cells(0, 0).Style.Font.Weight = ExcelFont.BoldWeight
        ws1.Cells(0, 1).Style.Font.Weight = ExcelFont.BoldWeight
        ws1.Columns(0).Width = CInt(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart))
        ws1.Columns(1).Style.NumberFormat = """$""#,##0"

        ' Create Excel chart sheet.
        Dim ws2 = ef.Worksheets.Add(SheetType.Chart, "ChartSheet")

        ' Create Excel chart and select data for it.
        ' You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        Dim chart = ws2.Charts.Add(ChartType.Bar, 0, 0, 10, 10, LengthUnit.Centimeter)
        chart.SelectData(ws1.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), True)

        ef.Save("Chart Sheet.xlsx")

    End Sub

End Module
