Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' Load Excel workbook from file's path.
        Dim workbook As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        ' Iterate through all worksheets in a workbook.
        For Each worksheet As ExcelWorksheet In workbook.Worksheets

            ' Display sheet's name.
            Console.WriteLine("{1} {0} {1}" & vbLf, worksheet.Name, New String("#"c, 30))

            ' Iterate through all rows in a worksheet.
            For Each row As ExcelRow In worksheet.Rows

                ' Iterate through all allocated cells in a row.
                For Each cell As ExcelCell In row.AllocatedCells

                    ' Read cell's data.
                    Dim value As String = If(cell.Value?.ToString(), "EMPTY")

                    ' Display cell's value and type.
                    value = If(value.Length > 15, value.Remove(15) & "…", value)
                    Console.Write($"{value} [{cell.ValueType}]".PadRight(30))
                Next

                Console.WriteLine()
            Next
        Next
    End Sub

    Sub Example2()
        Dim workbook As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        For sheetIndex As Integer = 0 To workbook.Worksheets.Count - 1

            ' Get Excel worksheet using zero-based index.
            Dim worksheet As ExcelWorksheet = workbook.Worksheets(sheetIndex)
            Console.WriteLine($"Sheet name: ""{worksheet.Name}""")
            Console.WriteLine($"Sheet index: {worksheet.Index}" & vbLf)

            For rowIndex As Integer = 0 To worksheet.Rows.Count - 1

                ' Get Excel row using zero-based index.
                Dim row As ExcelRow = worksheet.Rows(rowIndex)
                Console.WriteLine($"Row name: ""{row.Name}""")
                Console.WriteLine($"Row index: {row.Index}")

                Console.Write("Cell names:")
                For columnIndex As Integer = 0 To row.AllocatedCells.Count - 1

                    ' Get Excel cell using zero-based index.
                    Dim cell As ExcelCell = row.Cells(columnIndex)
                    Console.Write($" ""{cell.Name}"",")
                Next
                Console.WriteLine(vbLf)
            Next
        Next
    End Sub

    Sub Example3()
        Dim workbook As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        For Each worksheet As ExcelWorksheet In workbook.Worksheets

            Dim enumerator As CellRangeEnumerator = worksheet.Cells.GetReadEnumerator()
            While enumerator.MoveNext()

                Dim cell As ExcelCell = enumerator.Current
                Console.WriteLine($"Cell ""{cell.Name}"" [{cell.Row.Index}, {cell.Column.Index}]: {cell.Value}")
            End While
        Next
    End Sub

End Module