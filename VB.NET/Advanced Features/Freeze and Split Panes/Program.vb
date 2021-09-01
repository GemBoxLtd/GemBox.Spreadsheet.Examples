Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile

        ' Frozen Rows (first 2 rows are frozen).
        Dim worksheet1 = workbook.Worksheets.Add("Frozen rows")
        worksheet1.Panes = New WorksheetPanes(PanesState.Frozen, 0, 2, "A3", PanePosition.BottomLeft)

        ' Frozen Columns (first column is frozen).
        Dim worksheet2 = workbook.Worksheets.Add("Frozen columns")
        worksheet2.Panes = New WorksheetPanes(PanesState.Frozen, 1, 0, "B1", PanePosition.TopRight)

        ' Frozen Rows and Columns (first 2 rows and first 3 columns are frozen).
        Dim worksheet3 = workbook.Worksheets.Add("Frozen rows and columns")
        worksheet3.Panes = New WorksheetPanes(PanesState.Frozen, 3, 2, "E5", PanePosition.BottomRight)

        ' Split pane.
        Dim worksheet4 = workbook.Worksheets.Add("Split pane")
        worksheet4.Panes = New WorksheetPanes(PanesState.Split, 2310, 1500, "D7", PanePosition.BottomRight)

        workbook.Save("Freeze or Split Panes.xlsx")
    End Sub
End Module