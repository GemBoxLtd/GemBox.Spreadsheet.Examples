Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
    End Sub

    Sub Example1()
        Dim workbook = ExcelFile.Load("NumberFormat.xlsx")

        Dim worksheet = workbook.Worksheets(0)

        worksheet.Cells(0, 2).Value = "ExcelCell.Value"
        worksheet.Columns(2).Style.NumberFormat = "@"

        worksheet.Cells(0, 3).Value = "CellStyle.NumberFormat"
        worksheet.Columns(3).Style.NumberFormat = "@"

        worksheet.Cells(0, 4).Value = "ExcelCell.GetFormattedValue()"
        worksheet.Columns(4).Style.NumberFormat = "@"

        For i As Integer = 0 To worksheet.Rows.Count

            Dim sourceCell = worksheet.Cells(i, 0)

            worksheet.Cells(i, 2).Value = sourceCell.Value?.ToString()
            worksheet.Cells(i, 3).Value = sourceCell.Style.NumberFormat
            worksheet.Cells(i, 4).Value = sourceCell.GetFormattedValue()
        Next

        ' Set column widths.
        Dim columnWidths = New Double() {192, Double.NaN, 122, 236, 200}
        For i As Integer = 0 To columnWidths.Length - 1
            If Not Double.IsNaN(columnWidths(i)) Then worksheet.Columns(i).SetWidth(columnWidths(i), LengthUnit.Pixel)
        Next

        workbook.Save("Number Format.xlsx")
    End Sub

    Sub Example2()
        Dim workbook = New ExcelFile()

        Dim worksheet = workbook.Worksheets.Add("sheet")

        worksheet.Columns(0).SetWidth(200, LengthUnit.Pixel)

        ' Show the value as a number with two decimal places And thousands separator.
        worksheet.Cells(0, 0).Style.NumberFormat = NumberFormatBuilder.Number(2, useThousandsSeparator:=True)
        worksheet.Cells(0, 0).Value = 2500.333

        ' Show the value in Euros And display negative values in parentheses.
        worksheet.Cells(1, 0).Style.NumberFormat = NumberFormatBuilder.Currency("€", 2, useParenthesesToDisplayNegativeValue:=True)
        worksheet.Cells(1, 0).Value = -50

        ' Show the value in accounting format with three decimal places.
        worksheet.Cells(2, 0).Style.NumberFormat = NumberFormatBuilder.Accounting(3, currencySymbol:="$")
        worksheet.Cells(2, 0).Value = -50

        ' Show the value in ISO 8061 date format.
        worksheet.Cells(3, 0).Style.NumberFormat = NumberFormatBuilder.DateTimeIso8061()
        worksheet.Cells(3, 0).Value = DateTime.Now

        ' Show the value as percentage.
        worksheet.Cells(4, 0).Style.NumberFormat = NumberFormatBuilder.Percentage(2)
        worksheet.Cells(4, 0).Value = 1 / 3D

        ' Show the value as fraction with 100 as a denominator.
        worksheet.Cells(5, 0).Style.NumberFormat = NumberFormatBuilder.FractionWithPreciseDenominator(100)
        worksheet.Cells(5, 0).Value = 1 / 3D

        ' Show the value in scientific notation using two decimal places.
        worksheet.Cells(6, 0).Style.NumberFormat = NumberFormatBuilder.Scientific(2)
        worksheet.Cells(6, 0).Value = Math.Pow(Math.PI, 10)

        workbook.Save("Number Format Builder.docx")
    End Sub
End Module