Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Formula Utility Methods")

        ' Fill first column with values.
        Dim i As Int32
        For i = 0 To 9 Step 1
            ws.Cells(i, 0).Value = i + 1
        Next

        ' Cell B1 has formula '=A1*2', B2 '=A2*2', etc.
        For i = 0 To 9 Step 1
            ws.Cells(i, 1).Formula = String.Format("={0}*2", CellRange.RowColumnToPosition(i, 0))
        Next

        ' Cell C1 has formula '=SUM(A1:B1)', C2 '=SUM(A2:B2)', etc.
        For i = 0 To 9 Step 1
            ws.Cells(i, 2).Formula = String.Format("=SUM(A{0}:B{0})", ExcelRowCollection.RowIndexToName(i))
        Next

        ' Cell A12 contains sum of all values from the first row.
        ws.Cells("A12").Formula = String.Format("=SUM(A1:{0}1)", ExcelColumnCollection.ColumnIndexToName(ws.Rows(0).AllocatedCells.Count - 1))

        ef.Save("Formula Utility Methods.xlsx")

    End Sub

End Module