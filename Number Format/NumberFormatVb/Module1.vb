Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("NumberFormat.xlsx")

        Dim ws = ef.Worksheets(0)

        ws.Cells(0, 2).Value = "ExcelCell.Value"
        ws.Columns(2).Style.NumberFormat = "@"

        ws.Cells(0, 3).Value = "CellStyle.NumberFormat"
        ws.Columns(3).Style.NumberFormat = "@"

        ws.Cells(0, 4).Value = "ExcelCell.GetFormattedValue()"
        ws.Columns(4).Style.NumberFormat = "@"

        For i As Integer = 0 To ws.Rows.Count
            Dim sourceCell = ws.Cells(i, 0)

            If (sourceCell.Value IsNot Nothing) Then
                ws.Cells(i, 2).Value = sourceCell.Value.ToString()
            End If

            ws.Cells(i, 3).Value = sourceCell.Style.NumberFormat
            ws.Cells(i, 4).Value = sourceCell.GetFormattedValue()
        Next

        ' Auto-fit columns
        For i As Integer = 0 To 4
            ws.Columns(i).AutoFit()
        Next

        ef.Save("Number Format.xlsx")

    End Sub

End Module