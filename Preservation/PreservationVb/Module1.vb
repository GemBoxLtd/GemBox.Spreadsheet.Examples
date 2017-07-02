Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("ChartTemplate.xlsx")

        Dim numberOfEmployees As Integer = 4

        Dim ws = ef.Worksheets(0)

        ' Update named ranges 'Names' and 'Salaries' which are used by preserved chart.
        ws.NamedRanges("Names").Range = ws.Cells.GetSubrangeAbsolute(1, 0, numberOfEmployees, 0)
        ws.NamedRanges("Salaries").Range = ws.Cells.GetSubrangeAbsolute(1, 1, numberOfEmployees, 1)

        ' Add data which is used by preserved chart through named ranges 'Names' and 'Salaries'.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1
            ws.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))
            ws.Cells(i + 1, 1).SetValue(random.Next(1000, 5000))
        Next

        ef.Save("Preservation.xlsx")

    End Sub

End Module