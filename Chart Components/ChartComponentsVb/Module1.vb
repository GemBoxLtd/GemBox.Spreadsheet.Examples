Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfEmployees = 4
        Dim numberOfYears = 4

        Dim ef = New ExcelFile()
        Dim ws = ef.Worksheets.Add("Chart")

        ' Create chart and select data for it.
        Dim chart = DirectCast(ws.Charts.Add(ChartType.Column, "B7", "O27"), ColumnChart)
        chart.SelectData(ws.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, numberOfYears))

        ' Set chart title.
        chart.Title.Text = "Clustered Column Chart"

        ' Set axis titles.
        chart.Axes.Horizontal.Title.Text = "Years"
        chart.Axes.Vertical.Title.Text = "Salaries"

        ' For all charts (except Pie and Bar) value axis is vertical.
        Dim valueAxis = chart.Axes.VerticalValue

        ' Set value axis scaling, units, gridlines and tick marks.
        valueAxis.Minimum = 0
        valueAxis.Maximum = 6000
        valueAxis.MajorUnit = 1000
        valueAxis.MinorUnit = 500
        valueAxis.MajorGridlines.IsVisible = True
        valueAxis.MinorGridlines.IsVisible = True
        valueAxis.MajorTickMarkType = TickMarkType.Outside
        valueAxis.MinorTickMarkType = TickMarkType.Cross

        ' Add data which is used by the chart.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1
            ws.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))

            For j As Integer = 0 To numberOfYears - 1
                ws.Cells(i + 1, j + 1).SetValue(random.Next(1000, 5000))
            Next
        Next

        ' Set header row and formatting.
        ws.Cells(0, 0).Value = "Name"
        ws.Cells(0, 0).Style.Font.Weight = ExcelFont.BoldWeight
        ws.Columns(0).Width = CInt(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart))
        Dim startYear As Integer = DateTime.Now.Year - numberOfYears
        For i As Integer = 0 To numberOfYears - 1
            ws.Cells(0, i + 1).SetValue(startYear + i)
            ws.Cells(0, i + 1).Style.Font.Weight = ExcelFont.BoldWeight
            ws.Cells(0, i + 1).Style.NumberFormat = "General"
            ws.Columns(i + 1).Style.NumberFormat = """$""#,##0"
        Next

        ' Make entire sheet print horizontally centered on a single page with headings and gridlines.
        Dim printOptions = ws.PrintOptions
        printOptions.HorizontalCentered = True
        printOptions.PrintHeadings = True
        printOptions.PrintGridlines = True
        printOptions.FitWorksheetWidthToPages = 1
        printOptions.FitWorksheetHeightToPages = 1

        ef.Save("Chart Components.xlsx")

    End Sub

End Module