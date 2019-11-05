Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
End Module