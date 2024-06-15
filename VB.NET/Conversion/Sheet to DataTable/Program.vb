Imports GemBox.Spreadsheet
Imports System
Imports System.Data

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        ' Create DataTable with specified columns.
        Dim dataTable As New DataTable()
        dataTable.Columns.Add("First_Column", GetType(String))
        dataTable.Columns.Add("Second_Column", GetType(String))
        dataTable.Columns.Add("Third_Column", GetType(Integer))
        dataTable.Columns.Add("Fourth_Column", GetType(Double))

        ' Select the first worksheet from the file.
        Dim worksheet = workbook.Worksheets(0)

        ' Extract the data from an Excel worksheet to the DataTable.
        Dim options As New ExtractToDataTableOptions(0, 0, 20)
        AddHandler options.ExcelCellToDataTableCellConverting,
            Sub(sender, e)
                If Not e.IsDataTableValueValid Then
                    ' Convert ExcelCell value to string.
                    e.DataTableValue = If(e.DataTableColumnType = GetType(String),
                        e.ExcelCell.Value?.ToString(),
                        DBNull.Value)
                End If
            End Sub
        worksheet.ExtractToDataTable(dataTable, options)

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

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        ' Create DataTable with specified columns.
        Dim dataTable As New DataTable()
        dataTable.Columns.Add("First_Column", GetType(String))
        dataTable.Columns.Add("Second_Column", GetType(String))
        dataTable.Columns.Add("Third_Column", GetType(Integer))
        dataTable.Columns.Add("Fourth_Column", GetType(Double))

        ' Select the first worksheet from the file.
        Dim worksheet = workbook.Worksheets(0)

        ' Extract the data from an Excel worksheet to the DataTable.
        Dim options As New ExtractToDataTableOptions(0, 0, 20)
        AddHandler options.ExcelCellToDataTableCellConverting,
            Sub(sender, e)
                If Not e.IsDataTableValueValid Then
                    ' Convert ExcelCell value to string.
                    e.DataTableValue = If(e.DataTableColumnType = GetType(String),
                        e.ExcelCell.Value?.ToString(),
                        DBNull.Value)
                End If
            End Sub
        worksheet.ExtractToDataTable(dataTable, options)

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
