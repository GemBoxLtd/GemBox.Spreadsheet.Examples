Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Inline Text Formatting")

        worksheet.Cells(0, 0).Value = "Inline text formatting examples:"
        worksheet.PrintOptions.PrintGridlines = True

        ' Column width of 20 characters.
        worksheet.Columns(0).Width = 20 * 256

        worksheet.Cells(2, 0).Value = "This is big and red text!"

        ' Apply size to 'big and red' part of text
        worksheet.Cells(2, 0).GetCharacters(8, 11).Font.Size = 400

        ' Apply color to 'red' part of text
        worksheet.Cells(2, 0).GetCharacters(16, 3).Font.Color = SpreadsheetColor.FromName(ColorName.Red)

        ' Format cell content
        worksheet.Cells(4, 0).Value = "Formatting selected characters with GemBox.Spreadsheet component."
        worksheet.Cells(4, 0).Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
        worksheet.Cells(4, 0).Style.Font.Italic = True
        worksheet.Cells(4, 0).Style.WrapText = True

        ' Get characters from index 36 to the end of string
        Dim characters = worksheet.Cells(4, 0).GetCharacters(36)

        ' Apply color and underline to selected characters
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange)
        characters.Font.UnderlineStyle = UnderlineStyle.Single

        ' Write selected characters
        worksheet.Cells(6, 0).Value = "Selected characters: " + characters.Text

        workbook.Save("Inline Text Formatting.xlsx")
    End Sub
End Module
