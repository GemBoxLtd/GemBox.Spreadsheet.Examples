Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts

Module Program

    Sub Main()

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create new workbook.
        Dim workbook As New ExcelFile()

        ' Create new worksheet.
        Dim worksheet = workbook.Worksheets.Add("Sheet1")
        worksheet.PrintOptions.PrintHeadings = True
        worksheet.PrintOptions.PrintGridlines = True
        worksheet.PrintOptions.FitWorksheetWidthToPages = 1

        ' Add sample formatting.
        worksheet.Rows(0).Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Columns(0).SetWidth(80, LengthUnit.Pixel)
        worksheet.Columns(1).SetWidth(80, LengthUnit.Pixel)

        ' Add sample data.
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

        ' Add sample chart.
        Dim chart = worksheet.Charts.Add(ChartType.Bar, "B7", "J20")
        chart.SelectData(worksheet.Cells.GetSubrange("A1:B5"), True)

        ' Save spreadsheet in XLSX and PDF format.
        workbook.Save("output.xlsx")
        workbook.Save("output.pdf")
    End Sub
End Module
