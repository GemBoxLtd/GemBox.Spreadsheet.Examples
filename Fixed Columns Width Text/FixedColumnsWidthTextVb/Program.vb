Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Define columns width (for input file format).
        Dim loadOptions As New FixedWidthLoadOptions(
            New FixedWidthColumn(8),
            New FixedWidthColumn(8),
            New FixedWidthColumn(8))

        ' Load file.
        Dim workbook = ExcelFile.Load("FixedColumnsWidthText.prn", loadOptions)

        ' Modify file.
        workbook.Worksheets.ActiveWorksheet.GetUsedCellRange(True).Sort(False).By(1).Apply()

        ' Define columns width (for output file format).
        Dim saveOptions As New FixedWidthSaveOptions(
            New FixedWidthColumn(8),
            New FixedWidthColumn(8),
            New FixedWidthColumn(8))

        workbook.Save("Fixed Columns Width Text.prn", saveOptions)
    End Sub
End Module