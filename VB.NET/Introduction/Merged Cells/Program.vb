Imports GemBox.Spreadsheet
Imports System.Linq

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Sheet1")

        ' Get the cell range.
        Dim range = worksheet.Cells.GetSubrange("B2:E5")

        ' Merge cells in the current range.
        range.Merged = True

        ' Set the value of the merged range.
        range.Value = "Merged"

        ' Set the style of the merged range.
        range.Style.VerticalAlignment = VerticalAlignmentStyle.Center

        ' Set the style of the merged range using a cell within.
        worksheet.Cells("C3").Style.Borders _
            .SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Double)

        workbook.Save("Merged Cells.xlsx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Merged Cells.xlsx")
        Dim worksheet = workbook.Worksheets(0)

        ' Get the first merged range.
        Dim mergedRange = worksheet.Rows _
            .SelectMany(Function(row) row.AllocatedCells) _
            .Select(Function(cell) cell.MergedRange) _
            .FirstOrDefault(Function(range) range IsNot Nothing)

        If mergedRange <> Nothing Then
            ' Important, you cannot unmerge the ExcelCell.MergedRange property.
            ' In other words, the following is not allowed:  mergedRange.Merged = False

            ' Instead, you need to retrieve the same CellRange from the ExcelWorksheet and then unmerge it.
            ' This kind of implementation was chosen for performance reasons.
            worksheet.Cells.GetSubrange(mergedRange.Name).Merged = False

            worksheet.Cells(mergedRange.StartPosition).Value = "Unmerged"
        End If

        workbook.Save("Unmerged Cells.xlsx")
    End Sub

End Module
