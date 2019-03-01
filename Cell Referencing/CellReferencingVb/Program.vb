Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Cell Referencing")

        worksheet.Cells(0).Value = "Cell referencing examples:"

        worksheet.Cells("B2").Value = "Cell B2."
        worksheet.Cells(6, 0).Value = "Cell in row 7 and column A."

        worksheet.Rows(2).Cells(0).Value = "Cell in row 3 and column A."
        worksheet.Rows("4").Cells("B").Value = "Cell in row 4 and column B."

        worksheet.Columns(2).Cells(4).Value = "Cell in column C and row 5."
        worksheet.Columns("AA").Cells("6").Value = "Cell in AA column and row 6."

        ' Referencing Excel row's cell range.
        Dim cellRange = worksheet.Rows(7).Cells

        cellRange(0).Value = cellRange.IndexingMode.ToString()
        cellRange(3).Value = "D8"
        cellRange("B").Value = "B8"

        ' Referencing Excel column's cell range.
        cellRange = worksheet.Columns(7).Cells

        cellRange(0).Value = cellRange.IndexingMode.ToString()
        cellRange(2).Value = "H3"
        cellRange("5").Value = "H5"

        ' Referencing arbitrary Excel cell range.
        cellRange = worksheet.Cells.GetSubrange("I2", "L8")
        cellRange.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromArgb(0, 0, 128), LineStyle.Dashed)

        cellRange("J7").Value = cellRange.IndexingMode.ToString()
        cellRange(0, 0).Value = "I2"
        cellRange("J3").Value = "J3"
        cellRange(4).Value = "I3" ' Cell range width is 4 (I J K L).

        ' Set column widths and some print options (for better look when exporting to PDF, XPS and printing).
        Dim columnWidths = New Double() {175, 174, 174, 24, Double.NaN, Double.NaN, Double.NaN, 54, 19, 81}
        For i As Integer = 0 To columnWidths.Length - 1
            If Not Double.IsNaN(columnWidths(i)) Then worksheet.Columns(i).SetWidth(columnWidths(i), LengthUnit.Pixel)
        Next

        worksheet.PrintOptions.PrintGridlines = True
        worksheet.PrintOptions.PrintHeadings = True

        workbook.Save("Cell Referencing.xlsx")
    End Sub
End Module