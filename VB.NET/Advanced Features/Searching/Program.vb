Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")
        Dim worksheet = workbook.Worksheets.ActiveWorksheet

        Dim searchText = "Apollo"
        For Each cell In worksheet.Cells.FindAllText(searchText)
            Console.WriteLine($"Text was found in cell '{cell.Name}' (""{cell.StringValue}"").")
        Next

    End Sub
End Module