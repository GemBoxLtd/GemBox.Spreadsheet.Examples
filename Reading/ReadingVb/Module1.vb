Imports System
Imports System.Text
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim sb As New StringBuilder()

        ' Iterate through all worksheets in an Excel workbook.
        For Each sheet As ExcelWorksheet In ef.Worksheets
            sb.AppendLine()
            sb.AppendFormat("{0} {1} {0}", New String("-"c, 25), sheet.Name)

            ' Iterate through all rows in an Excel worksheet.
            For Each row As ExcelRow In sheet.Rows
                sb.AppendLine()

                ' Iterate through all allocated cells in an Excel row.
                For Each cell As ExcelCell In row.AllocatedCells
                    If cell.ValueType <> CellValueType.Null Then
                        sb.Append(String.Format("{0} [{1}]", cell.Value, cell.ValueType).PadRight(25))
                    Else
                        sb.Append(New String(" "c, 25))
                    End If
                Next
            Next
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module