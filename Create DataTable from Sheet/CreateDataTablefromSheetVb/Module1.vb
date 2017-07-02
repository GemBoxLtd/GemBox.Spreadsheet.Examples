Imports System
Imports System.Data
Imports System.Text
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        ' Select the first worksheet from the file.
        Dim ws = ef.Worksheets(0)

        ' Create DataTable from an Excel worksheet.
        Dim dataTable As DataTable = ws.CreateDataTable(New CreateDataTableOptions() With
        {
            .ColumnHeaders = True,
            .StartRow = 1,
            .NumberOfColumns = 5,
            .NumberOfRows = ws.Rows.Count - 1,
            .Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        })

        ' Write DataTable content
        Dim sb = New StringBuilder()
        sb.AppendLine("DataTable content:")
        For Each row As DataRow In dataTable.Rows
            sb.AppendFormat("{0}" & vbTab & "{1}" & vbTab & "{2}" & vbTab & "{3}" & vbTab & "{4}", row(0), row(1), row(2), row(3), row(4))
            sb.AppendLine()
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module