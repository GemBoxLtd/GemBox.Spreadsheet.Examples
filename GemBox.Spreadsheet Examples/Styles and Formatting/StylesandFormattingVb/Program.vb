Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Styles and Formatting")

        worksheet.Cells(0, 1).Value = "Cell style examples:"
        worksheet.PrintOptions.PrintGridlines = True

        Dim row As Integer = 0

        ' Column width of 4, 30 and 36 characters.
        worksheet.Columns(0).Width = 4 * 256
        worksheet.Columns(1).Width = 30 * 256
        worksheet.Columns(2).Width = 36 * 256

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Borders.SetBorders(...)"
        worksheet.Cells(row, 2).Style.Borders.SetBorders(MultipleBorders.All Or MultipleBorders.Diagonal, SpreadsheetColor.FromArgb(252, 1, 1), LineStyle.Thin)

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.FillPattern.SetPattern(...)"
        worksheet.Cells(row, 2).Style.FillPattern.SetPattern(FillPatternStyle.ThinHorizontalCrosshatch, SpreadsheetColor.FromName(ColorName.Green), SpreadsheetColor.FromName(ColorName.Yellow))

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Color ="
        worksheet.Cells(row, 2).Value = "Color.Blue"
        worksheet.Cells(row, 2).Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Italic ="
        worksheet.Cells(row, 2).Value = "true"
        worksheet.Cells(row, 2).Style.Font.Italic = True

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Name ="
        worksheet.Cells(row, 2).Value = "Comic Sans MS"
        worksheet.Cells(row, 2).Style.Font.Name = "Comic Sans MS"

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.ScriptPosition ="
        worksheet.Cells(row, 2).Value = "ScriptPosition.Superscript"
        worksheet.Cells(row, 2).Style.Font.ScriptPosition = ScriptPosition.Superscript

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Size ="
        worksheet.Cells(row, 2).Value = "18 * 20"
        worksheet.Cells(row, 2).Style.Font.Size = 18 * 20

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Strikeout ="
        worksheet.Cells(row, 2).Value = "true"
        worksheet.Cells(row, 2).Style.Font.Strikeout = True

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.UnderlineStyle ="
        worksheet.Cells(row, 2).Value = "UnderlineStyle.Double"
        worksheet.Cells(row, 2).Style.Font.UnderlineStyle = UnderlineStyle.Double

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Font.Weight ="
        worksheet.Cells(row, 2).Value = "ExcelFont.BoldWeight"
        worksheet.Cells(row, 2).Style.Font.Weight = ExcelFont.BoldWeight

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.HorizontalAlignment ="
        worksheet.Cells(row, 2).Value = "HorizontalAlignmentStyle.Center"
        worksheet.Cells(row, 2).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Indent"
        worksheet.Cells(row, 2).Value = "five"
        worksheet.Cells(row, 2).Style.HorizontalAlignment = HorizontalAlignmentStyle.Left
        worksheet.Cells(row, 2).Style.Indent = 5

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.IsTextVertical = "
        worksheet.Cells(row, 2).Value = "true"
        ' Set row height to 60 points.
        worksheet.Rows(row).Height = 60 * 20
        worksheet.Cells(row, 2).Style.IsTextVertical = True

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.NumberFormat"
        worksheet.Cells(row, 2).Value = 1234
        worksheet.Cells(row, 2).Style.NumberFormat = "#.##0,00 [$Krakozhian Money Units]"

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.Rotation"
        worksheet.Cells(row, 2).Value = "35 degrees up"
        worksheet.Cells(row, 2).Style.Rotation = 35

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.ShrinkToFit"
        worksheet.Cells(row, 2).Value = "This property is set to true so this text appears shrunk."
        worksheet.Cells(row, 2).Style.ShrinkToFit = True

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.VerticalAlignment ="
        worksheet.Cells(row, 2).Value = "VerticalAlignmentStyle.Top"
        ' Set row height to 30 points.
        worksheet.Rows(row).Height = 30 * 20
        worksheet.Cells(row, 2).Style.VerticalAlignment = VerticalAlignmentStyle.Top

        row = row + 2
        worksheet.Cells(row, 1).Value = ".Style.WrapText"
        worksheet.Cells(row, 2).Value = "This property is set to true so this text appears broken into multiple lines."
        worksheet.Cells(row, 2).Style.WrapText = True

        workbook.Save("Styles and Formatting.xlsx")
    End Sub
End Module