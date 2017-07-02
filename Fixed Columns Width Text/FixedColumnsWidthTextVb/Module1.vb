Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Define columns width (for input file format)
        Dim loadOptions As New FixedWidthLoadOptions(
            New FixedWidthColumn(8),
            New FixedWidthColumn(8),
            New FixedWidthColumn(8))

        ' Load file
        Dim ef As ExcelFile = ExcelFile.Load("FixedColumnsWidthText.prn", loadOptions)

        ' Modify file
        ef.Worksheets.ActiveWorksheet.GetUsedCellRange(True).Sort(False).By(1).Apply()

        ' Define columns width (for output file format)
        Dim saveOptions As New FixedWidthSaveOptions(
            New FixedWidthColumn(8),
            New FixedWidthColumn(8),
            New FixedWidthColumn(8))

        ef.Save("FixedColumnsWidthText.prn", saveOptions)

    End Sub

End Module