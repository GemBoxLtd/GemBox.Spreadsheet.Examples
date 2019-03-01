Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfEmployees = 4
        Dim numberOfYears = 4

        Dim workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Chart")

        ' Create chart and select data for it.
        Dim chart = DirectCast(worksheet.Charts.Add(ChartType.Column, "B7", "O27"), ColumnChart)
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, numberOfYears))

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
            worksheet.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))

            For j As Integer = 0 To numberOfYears - 1
                worksheet.Cells(i + 1, j + 1).SetValue(random.Next(1000, 5000))
            Next
        Next

        ' Set header row and formatting.
        worksheet.Cells(0, 0).Value = "Name"
        worksheet.Cells(0, 0).Style.Font.Weight = ExcelFont.BoldWeight
        worksheet.Columns(0).Width = CInt(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart))
        Dim startYear As Integer = DateTime.Now.Year - numberOfYears
        For i As Integer = 0 To numberOfYears - 1
            worksheet.Cells(0, i + 1).SetValue(startYear + i)
            worksheet.Cells(0, i + 1).Style.Font.Weight = ExcelFont.BoldWeight
            worksheet.Cells(0, i + 1).Style.NumberFormat = "General"
            worksheet.Columns(i + 1).Style.NumberFormat = """$""#,##0"
        Next

        ' Make entire sheet print horizontally centered on a single page with headings and gridlines.
        Dim printOptions = worksheet.PrintOptions
        printOptions.HorizontalCentered = True
        printOptions.PrintHeadings = True
        printOptions.PrintGridlines = True
        printOptions.FitWorksheetWidthToPages = 1
        printOptions.FitWorksheetHeightToPages = 1

        workbook.Save("Chart Components.xlsx")
    End Sub
End Module