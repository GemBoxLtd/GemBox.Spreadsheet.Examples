Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Chart")

        Dim numberOfEmployees As Integer = 4

        ' Create Excel chart and select data for it.
        Dim chart = ws.Charts.Add(ChartType.Bar, "D2", "M25")
        chart.SelectData(ws.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), True)

        ' Add data which is used by the Excel chart.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1
            ws.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))
            ws.Cells(i + 1, 1).SetValue(random.Next(1000, 5000))
        Next

        ' Set header row and formatting.
        ws.Cells(0, 0).Value = "Name"
        ws.Cells(0, 1).Value = "Salary"
        ws.Cells(0, 0).Style.Font.Weight = ExcelFont.BoldWeight
        ws.Cells(0, 1).Style.Font.Weight = ExcelFont.BoldWeight
        ws.Columns(0).Width = CInt(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart))
        ws.Columns(1).Style.NumberFormat = """$""#,##0"

        ' Make entire sheet print on a single page.
        ws.PrintOptions.FitWorksheetWidthToPages = 1
        ws.PrintOptions.FitWorksheetHeightToPages = 1

        ef.Save("Chart.xlsx")

    End Sub

End Module