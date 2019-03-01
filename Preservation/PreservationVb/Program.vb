Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("ChartTemplate.xlsx")

        Dim numberOfEmployees As Integer = 4

        Dim worksheet = workbook.Worksheets(0)

        ' Update named ranges 'Names' and 'Salaries' which are used by preserved chart.
        worksheet.NamedRanges("Names").Range = worksheet.Cells.GetSubrangeAbsolute(1, 0, numberOfEmployees, 0)
        worksheet.NamedRanges("Salaries").Range = worksheet.Cells.GetSubrangeAbsolute(1, 1, numberOfEmployees, 1)

        ' Add data which is used by preserved chart through named ranges 'Names' and 'Salaries'.
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        Dim random = New Random()
        For i As Integer = 0 To numberOfEmployees - 1

            worksheet.Cells(i + 1, 0).Value = names(i Mod names.Length) & (If(i < names.Length, String.Empty, " "c & (i \ names.Length + 1).ToString()))
            worksheet.Cells(i + 1, 1).SetValue(random.Next(1000, 5000))
        Next

        workbook.Save("Preservation.xlsx")
    End Sub
End Module