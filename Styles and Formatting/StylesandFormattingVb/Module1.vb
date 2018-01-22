Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Styles and Formatting")

        ws.Cells(0, 1).Value = "Cell style examples:"
        ws.PrintOptions.PrintGridlines = True

        Dim row As Integer = 0

        ' Column width of 4, 30 and 36 characters.
        ws.Columns(0).Width = 4 * 256
        ws.Columns(1).Width = 30 * 256
        ws.Columns(2).Width = 36 * 256

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Borders.SetBorders(...)"
        ws.Cells(row, 2).Style.Borders.SetBorders(MultipleBorders.All Or MultipleBorders.Diagonal, SpreadsheetColor.FromArgb(252, 1, 1), LineStyle.Thin)

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.FillPattern.SetPattern(...)"
        ws.Cells(row, 2).Style.FillPattern.SetPattern(FillPatternStyle.ThinHorizontalCrosshatch, SpreadsheetColor.FromName(ColorName.Green), SpreadsheetColor.FromName(ColorName.Yellow))

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Color ="
        ws.Cells(row, 2).Value = "Color.Blue"
        ws.Cells(row, 2).Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Italic ="
        ws.Cells(row, 2).Value = "true"
        ws.Cells(row, 2).Style.Font.Italic = True

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Name ="
        ws.Cells(row, 2).Value = "Comic Sans MS"
        ws.Cells(row, 2).Style.Font.Name = "Comic Sans MS"

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.ScriptPosition ="
        ws.Cells(row, 2).Value = "ScriptPosition.Superscript"
        ws.Cells(row, 2).Style.Font.ScriptPosition = ScriptPosition.Superscript

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Size ="
        ws.Cells(row, 2).Value = "18 * 20"
        ws.Cells(row, 2).Style.Font.Size = 18 * 20

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Strikeout ="
        ws.Cells(row, 2).Value = "true"
        ws.Cells(row, 2).Style.Font.Strikeout = True

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.UnderlineStyle ="
        ws.Cells(row, 2).Value = "UnderlineStyle.Double"
        ws.Cells(row, 2).Style.Font.UnderlineStyle = UnderlineStyle.Double

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Font.Weight ="
        ws.Cells(row, 2).Value = "ExcelFont.BoldWeight"
        ws.Cells(row, 2).Style.Font.Weight = ExcelFont.BoldWeight

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.HorizontalAlignment ="
        ws.Cells(row, 2).Value = "HorizontalAlignmentStyle.Center"
        ws.Cells(row, 2).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Indent"
        ws.Cells(row, 2).Value = "five"
        ws.Cells(row, 2).Style.HorizontalAlignment = HorizontalAlignmentStyle.Left
        ws.Cells(row, 2).Style.Indent = 5

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.IsTextVertical = "
        ws.Cells(row, 2).Value = "true"
        ' Set row height to 60 points.
        ws.Rows(row).Height = 60 * 20
        ws.Cells(row, 2).Style.IsTextVertical = True

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.NumberFormat"
        ws.Cells(row, 2).Value = 1234
        ws.Cells(row, 2).Style.NumberFormat = "#.##0,00 [$Krakozhian Money Units]"

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.Rotation"
        ws.Cells(row, 2).Value = "35 degrees up"
        ws.Cells(row, 2).Style.Rotation = 35

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.ShrinkToFit"
        ws.Cells(row, 2).Value = "This property is set to true so this text appears shrunk."
        ws.Cells(row, 2).Style.ShrinkToFit = True

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.VerticalAlignment ="
        ws.Cells(row, 2).Value = "VerticalAlignmentStyle.Top"
        ' Set row height to 30 points.
        ws.Rows(row).Height = 30 * 20
        ws.Cells(row, 2).Style.VerticalAlignment = VerticalAlignmentStyle.Top

        row = row + 2
        ws.Cells(row, 1).Value = ".Style.WrapText"
        ws.Cells(row, 2).Value = "This property is set to true so this text appears broken into multiple lines."
        ws.Cells(row, 2).Style.WrapText = True

        ef.Save("Styles and Formatting.xlsx")

    End Sub

End Module