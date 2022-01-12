Imports System
Imports System.Data
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        ' Select the first worksheet from the file.
        Dim worksheet = workbook.Worksheets(0)

        ' Create DataTable from an Excel worksheet.
        Dim dataTable = worksheet.CreateDataTable(New CreateDataTableOptions() With
        {
            .ColumnHeaders = True,
            .StartRow = 1,
            .NumberOfColumns = 5,
            .NumberOfRows = worksheet.Rows.Count - 1,
            .Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        })

        ' Write DataTable columns.
        For Each column As DataColumn In dataTable.Columns
            Console.Write(column.ColumnName.PadRight(20))
        Next
        Console.WriteLine()
        For Each column As DataColumn In dataTable.Columns
            Console.Write($"[{column.DataType}]".PadRight(20))
        Next
        Console.WriteLine()
        For Each column As DataColumn In dataTable.Columns
            Console.Write(New String("-"c, column.ColumnName.Length).PadRight(20))
        Next
        Console.WriteLine()

        ' Write DataTable rows.
        For Each row As DataRow In dataTable.Rows
            For Each item In row.ItemArray
                Dim value As String = item.ToString()
                value = If(value.Length > 20, value.Remove(19) & "…", value)
                Console.Write(value.PadRight(20))
            Next
            Console.WriteLine()
        Next

    End Sub
End Module