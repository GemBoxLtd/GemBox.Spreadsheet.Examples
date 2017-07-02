Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim ws = ef.Worksheets(0)

        Dim columnCount = ws.CalculateMaxUsedColumns()
        For i As Integer = 0 To columnCount - 1
            ws.Columns(i).AutoFit(1, ws.Rows(1), ws.Rows(ws.Rows.Count - 1))
        Next

        ef.Save("Row_Column AutoFit.pdf")

    End Sub

End Module