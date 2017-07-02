Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Sorting")

        Dim rnd = New Random()

        ws.Cells(0, 0).Value = "Sorted numbers"
        For i = 1 To 10 Step 1
            ws.Cells(i, 0).SetValue(rnd.Next(1, 100))
        Next

        ws.Cells.GetSubrangeAbsolute(1, 0, 10, 0).Sort(False).By(0).Apply()

        ws.Cells(0, 2).Value = "Sorted strings"
        ws.Cells(1, 2).Value = "John"
        ws.Cells(2, 2).Value = "Jennifer"
        ws.Cells(3, 2).Value = "Toby"
        ws.Cells(4, 2).Value = "Chloe"

        ws.Cells.GetSubrangeAbsolute(1, 2, 4, 2).Sort(False).By(0).Apply()

        ws.Cells(0, 4).Value = "Sorted by column E and after that by column F"
        For i = 1 To 10 Step 1
            ws.Cells(i, 4).SetValue(rnd.Next(1, 4))
            ws.Cells(i, 5).SetValue(rnd.Next(0, 10))
        Next

        ' Sort by column E ascending and then by column F descending.
        ' These sort settings will be saved to output XLSX file because they are active (parameter value is True).
        ws.Cells.GetSubrangeAbsolute(1, 4, 10, 5).Sort(True).By(0).By(1, True).Apply()

        ef.Save("Sorting.xlsx")

    End Sub

End Module