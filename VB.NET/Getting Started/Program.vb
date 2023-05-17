Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet As ExcelWorksheet = workbook.Worksheets.Add("Sheet1")
        Dim cell As ExcelCell = worksheet.Cells("A1")

        cell.Value = "Hello World!"

        workbook.Save("HelloWorld.xlsx")

    End Sub

End Module