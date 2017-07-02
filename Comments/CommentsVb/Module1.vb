Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Comments")

        ws.Cells.Item(0, 0).Value = "Comment examples:"

        ws.Cells.Item(2, 1).Comment.Text = "Empty cell."

        ws.Cells.Item(4, 1).Value = 5
        ws.Cells.Item(4, 1).Comment.Text = "Cell with a number."

        ws.Cells.Item("B7").Value = "Cell B7"

        Dim comment As ExcelComment = ws.Cells("B7").Comment
        comment.Text = "Some formatted text. Comment is:" & ChrW(10) & "a) multiline," & ChrW(10) & "b) large," & ChrW(10) & "c) visible, and " & ChrW(10) & "d) formatted."
        comment.IsVisible = True
        comment.TopLeftCell = New AnchorCell(ws.Columns(3), ws.Rows(4), True)
        comment.BottomRightCell = New AnchorCell(ws.Columns(5), ws.Rows(10), False)

        ' Get first 20 characters of a string
        Dim characters As FormattedCharacterRange = comment.GetCharacters(0, 20)

        ' Apply color, italic and size to selected characters
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange)
        characters.Font.Italic = True
        characters.Font.Size = 300

        ' Apply color to 'formatted' part of text
        comment.GetCharacters(5, 9).Font.Color = SpreadsheetColor.FromName(ColorName.Red)

        ef.Save("Comments.xlsx")

    End Sub

End Module