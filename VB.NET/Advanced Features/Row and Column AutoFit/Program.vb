Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim worksheet = workbook.Worksheets.ActiveWorksheet

        Dim columnCount = worksheet.CalculateMaxUsedColumns()
        For i As Integer = 0 To columnCount - 1
            worksheet.Columns(i).AutoFit(1, worksheet.Rows(1), worksheet.Rows(worksheet.Rows.Count - 1))
        Next

        workbook.Save("Row_Column AutoFit.xlsx")
    End Sub
End Module