Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Comments")

        worksheet.Cells.Item(0, 0).Value = "Comment examples:"

        worksheet.Cells.Item(2, 1).Comment.Text = "Empty cell."

        worksheet.Cells.Item(4, 1).Value = 5
        worksheet.Cells.Item(4, 1).Comment.Text = "Cell with a number."

        worksheet.Cells.Item("B7").Value = "Cell B7"

        Dim comment = worksheet.Cells("B7").Comment
        comment.Text = "Some formatted text. Comment is:" & ChrW(10) & "a) multiline," & ChrW(10) & "b) large," & ChrW(10) & "c) visible, and " & ChrW(10) & "d) formatted."
        comment.IsVisible = True
        comment.TopLeftCell = New AnchorCell(worksheet.Columns(3), worksheet.Rows(4), True)
        comment.BottomRightCell = New AnchorCell(worksheet.Columns(5), worksheet.Rows(10), False)

        ' Get first 20 characters of a string.
        Dim characters = comment.GetCharacters(0, 20)

        ' Apply color, italic and size to selected characters.
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange)
        characters.Font.Italic = True
        characters.Font.Size = 300

        ' Apply color to 'formatted' part of text.
        comment.GetCharacters(5, 9).Font.Color = SpreadsheetColor.FromName(ColorName.Red)

        workbook.Save("Comments.xlsx")
    End Sub
End Module