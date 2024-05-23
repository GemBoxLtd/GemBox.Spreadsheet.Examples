Imports GemBox.Spreadsheet
Imports System

Module Program

    Sub Main()

        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Formats")

        worksheet.Rows(0).Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Columns(0).Width = 25 * 256
        worksheet.Columns(1).Width = 35 * 256
        worksheet.Columns(2).Width = 25 * 256

        worksheet.Cells(0, 0).Value = "Value & Format"
        worksheet.Cells(0, 1).Value = "Format"
        worksheet.Cells(0, 2).Value = "Type"

        ' Sample data with values and formats.
        Dim data = New(Value As Object, Format As String)() _
        {
            (1.23, "0"),
            (1.23, "0.00"),
            (1.2345, "0.000"),
            (-2.345, "0.00_);[Red]\(0.00\)"),
            (2.34, "\$#,##0.00"),
            (2345.67, "#,##0.00\ [$�-1]"),
            (New DateTime(2012, 11, 9, 0, 0, 0), "[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"),
            (New DateTime(2012, 12, 5, 0, 0, 0), "[$-409]mmmm\ d\,\ yyyy;@"),
            (New DateTime(2012, 8, 10, 0, 0, 0), "yyyy\-mm\-dd\ \(dddd\)"),
            (New DateTime(2012, 8, 12, 0, 13, 0), "[$-409]m/d/yy\ h:mm\ AM/PM;@"),
            (New DateTime(2012, 8, 1, 21, 10, 0), "[$-409]h:mm\ AM/PM;@"),
            (New DateTime(1900, 1, 1, 6, 45, 30), "[h]:mm:ss"),
            (0.0123, "0%"),
            (0.0123, "0.00%"),
            (120000, "0.00E+00"),
            (1.25, "# ?/?"),
            (1.25, "#\ ?/100"),
            ("Sample text", "@")
        }

        For i = 0 To data.Length - 1
            Dim item = data(i)

            ' Write value and set number format to a cell.
            worksheet.Cells(i + 1, 0).Value = item.Value
            worksheet.Cells(i + 1, 0).Style.NumberFormat = item.Format

            ' Write number format as cell's value.
            worksheet.Cells(i + 1, 1).Value = item.Format

            ' Write data type as cell's value.
            worksheet.Cells(i + 1, 2).Value = item.Value.GetType().ToString()
        Next

        workbook.Save("Number Formats.xlsx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("sheet")
        worksheet.Columns(0).SetWidth(200, LengthUnit.Pixel)

        ' Show the value as a number with two decimal places And thousands separator.
        worksheet.Cells(0, 0).Style.NumberFormat =
            NumberFormatBuilder.Number(2, useThousandsSeparator:=True)
        worksheet.Cells(0, 0).Value = 2500.333

        ' Show the value in Euros And display negative values in parentheses.
        worksheet.Cells(1, 0).Style.NumberFormat =
            NumberFormatBuilder.Currency("�", 2, useParenthesesToDisplayNegativeValue:=True)
        worksheet.Cells(1, 0).Value = -50

        ' Show the value in accounting format with three decimal places.
        worksheet.Cells(2, 0).Style.NumberFormat =
            NumberFormatBuilder.Accounting(3, currencySymbol:="$")
        worksheet.Cells(2, 0).Value = -50

        ' Show the value in ISO 8061 date format.
        worksheet.Cells(3, 0).Style.NumberFormat =
            NumberFormatBuilder.DateTimeIso8061()
        worksheet.Cells(3, 0).Value = DateTime.Now

        ' Show the value as percentage.
        worksheet.Cells(4, 0).Style.NumberFormat =
            NumberFormatBuilder.Percentage(2)
        worksheet.Cells(4, 0).Value = 1 / 3D

        ' Show the value as fraction with 100 as a denominator.
        worksheet.Cells(5, 0).Style.NumberFormat =
            NumberFormatBuilder.FractionWithPreciseDenominator(100)
        worksheet.Cells(5, 0).Value = 1 / 3D

        ' Show the value in scientific notation using two decimal places.
        worksheet.Cells(6, 0).Style.NumberFormat =
            NumberFormatBuilder.Scientific(2)
        worksheet.Cells(6, 0).Value = Math.Pow(Math.PI, 10)

        workbook.Save("Number Format Builder.xlsx")
    End Sub
End Module
