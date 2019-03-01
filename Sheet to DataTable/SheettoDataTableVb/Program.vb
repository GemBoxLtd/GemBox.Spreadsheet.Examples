Imports System
Imports System.Data
Imports System.Text
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim dataTable = New DataTable

        ' Depending on the format of the input file, you need to change this:
        dataTable.Columns.Add("FirstColumn", GetType(String))
        dataTable.Columns.Add("SecondColumn", GetType(String))

        ' Select the first worksheet from the file.
        Dim worksheet = workbook.Worksheets(0)

        Dim options = New ExtractToDataTableOptions(0, 0, 10)
        options.ExtractDataOptions = ExtractDataOptions.StopAtFirstEmptyRow
        AddHandler options.ExcelCellToDataTableCellConverting,
            Sub(sender, e)
                If Not e.IsDataTableValueValid Then

                    ' GemBox.Spreadsheet doesn't automatically convert numbers to strings in ExtractToDataTable() method because of culture issues; 
                    ' someone would expect the number 12.4 as "12.4" and someone else as "12,4".
                    e.DataTableValue = If(e.ExcelCell.Value Is Nothing, Nothing, e.ExcelCell.Value.ToString())
                    e.Action = ExtractDataEventAction.Continue
                End If
            End Sub

        ' Extract the data from an Excel worksheet to the DataTable.
        ' Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
        worksheet.ExtractToDataTable(dataTable, options)

        ' Write DataTable content.
        Dim sb = New StringBuilder
        sb.AppendLine("DataTable content:")
        For Each row As DataRow In dataTable.Rows

            sb.AppendFormat("{0}    {1}", row(0), row(1))
            sb.AppendLine()
        Next

        Console.WriteLine(sb.ToString())
    End Sub
End Module