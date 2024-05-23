Imports System
Imports System.Text.RegularExpressions
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")
        Dim worksheet = workbook.Worksheets.ActiveWorksheet

        ' Find first cell with specific text.
        Dim searchText As String = "Ranger"
        Dim foundCell As ExcelCell = Nothing
        If worksheet.Cells.FindText(searchText, foundCell) Then
            Console.WriteLine($"First cell with '{searchText}' text:")
            Console.WriteLine($"Name: {foundCell.Name} | Value: ""{foundCell.StringValue}""")
            Console.WriteLine()
        End If

        ' Find all cells with specific text.
        searchText = "Apollo"
        Console.WriteLine($"All cells with '{searchText}' text:")

        For Each cell In worksheet.Cells.FindAllText(searchText)
            Console.WriteLine($"Name: {cell.Name} | Value: ""{cell.StringValue}""")
        Next
    End Sub

    Sub Example2()
        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")
        Dim worksheet = workbook.Worksheets.ActiveWorksheet

        ' Replace specific text in first cell in which it occurs.
        Dim searchText = "Ranger"
        Dim foundCell As ExcelCell = Nothing
        If worksheet.Cells.FindText(searchText, foundCell) Then
            foundCell.ReplaceText(searchText, "REPLACED FIRST")
        End If

        ' Replace specific text in all cells in which it occurs.
        worksheet.Cells.ReplaceText("Apollo", "REPLACED ALL")

        ' Replace specific regex pattern in all cells in which it occurs.
        Dim searchRegex = New Regex("Luna (\d{2})")
        worksheet.Cells.ReplaceText(searchRegex, "REPLACED $1")

        workbook.Save("FoundAndReplaced.xlsx")
    End Sub

End Module