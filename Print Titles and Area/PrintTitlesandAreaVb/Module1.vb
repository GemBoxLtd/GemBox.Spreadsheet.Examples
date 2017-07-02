Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile

        ' Always print 1st row.
        Dim ws1 = ef.Worksheets.Add("Sheet1")
        ws1.NamedRanges.SetPrintTitles(ws1.Rows(0), 1)

        ' Set print area (from A1 to I120):
        ws1.NamedRanges.SetPrintArea(ws1.Cells.GetSubrange("A1", "I120"))

        ' Always print columns from A to F.
        Dim ws2 = ef.Worksheets.Add("Sheet2")
        ws2.NamedRanges.SetPrintTitles(ws2.Columns(0), 6)

        ' Always print columns from A to F and first row.
        Dim ws3 = ef.Worksheets.Add("Sheet3")
        ws3.NamedRanges.SetPrintTitles(ws3.Rows(0), 1, ws3.Columns(0), 6)

        ' Fill Sheet1 with some data
        For i As Integer = 0 To 8
            ws1.Cells(0, i).Value = "Column " + ExcelColumnCollection.ColumnIndexToName(i)
        Next

        For i As Integer = 1 To 119
            For j As Integer = 0 To 8
                ws1.Cells(i, j).SetValue(i + j)
            Next
        Next

        ef.Save("Print Titles and Area.xlsx")

    End Sub

End Module