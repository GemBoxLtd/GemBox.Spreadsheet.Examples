Imports System
Imports System.Text
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim sb = New StringBuilder()

        ' Iterate through all worksheets in an Excel workbook.
        For Each worksheet In workbook.Worksheets

            sb.AppendLine()
            sb.AppendFormat("{0} {1} {0}", New String("-"c, 25), worksheet.Name)

            ' Iterate through all rows in an Excel worksheet.
            For Each row In worksheet.Rows

                sb.AppendLine()

                ' Iterate through all allocated cells in an Excel row.
                For Each cell In row.AllocatedCells
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