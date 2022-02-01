Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")
        Dim worksheet = workbook.Worksheets(0)

        Dim searchText = "Apollo"
        Dim range = worksheet.Columns(0).Cells

        Dim row As Integer, column As Integer
        While range.FindText(searchText, row, column)

            Dim cell = worksheet.Cells(row, column)
            Console.WriteLine($"Text was found in cell '{cell.Name}' (""{cell.StringValue}"").")

            range = range.GetSubrangeAbsolute(row + 1, 0, worksheet.Rows.Count, 0)
        End While

    End Sub
End Module