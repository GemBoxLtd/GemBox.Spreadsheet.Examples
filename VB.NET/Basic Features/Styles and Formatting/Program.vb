Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Styles")

        worksheet.Rows(0).Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Columns(0).Width = 30 * 256
        worksheet.Columns(1).Width = 35 * 256

        worksheet.Cells(0, 0).Value = "Property"
        worksheet.Cells(0, 1).Value = "Result"

        Dim row As Integer = 2
        worksheet.Cells(row, 0).Value = "Borders"
        worksheet.Cells(row, 1).Style.Borders.SetBorders(
            MultipleBorders.All Or MultipleBorders.Diagonal,
            SpreadsheetColor.FromArgb(252, 1, 1),
            LineStyle.Thin)

        row = row + 2
        worksheet.Cells(row, 0).Value = "FillPattern"
        worksheet.Cells(row, 1).Style.FillPattern.SetPattern(
            FillPatternStyle.ThinHorizontalCrosshatch,
            SpreadsheetColor.FromName(ColorName.Green),
            SpreadsheetColor.FromName(ColorName.Yellow))

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Color"
        worksheet.Cells(row, 1).Value = "Color.Blue"
        worksheet.Cells(row, 1).Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Italic"
        worksheet.Cells(row, 1).Value = "true"
        worksheet.Cells(row, 1).Style.Font.Italic = True

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Name"
        worksheet.Cells(row, 1).Value = "Comic Sans MS"
        worksheet.Cells(row, 1).Style.Font.Name = "Comic Sans MS"

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.ScriptPosition"
        worksheet.Cells(row, 1).Value = "ScriptPosition.Superscript"
        worksheet.Cells(row, 1).Style.Font.ScriptPosition = ScriptPosition.Superscript

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Size"
        worksheet.Cells(row, 1).Value = "18 * 20"
        worksheet.Cells(row, 1).Style.Font.Size = 18 * 20

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Strikeout"
        worksheet.Cells(row, 1).Value = "true"
        worksheet.Cells(row, 1).Style.Font.Strikeout = True

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.UnderlineStyle"
        worksheet.Cells(row, 1).Value = "UnderlineStyle.Double"
        worksheet.Cells(row, 1).Style.Font.UnderlineStyle = UnderlineStyle.Double

        row = row + 2
        worksheet.Cells(row, 0).Value = "Font.Weight"
        worksheet.Cells(row, 1).Value = "ExcelFont.BoldWeight"
        worksheet.Cells(row, 1).Style.Font.Weight = ExcelFont.BoldWeight

        row = row + 2
        worksheet.Cells(row, 0).Value = "HorizontalAlignment"
        worksheet.Cells(row, 1).Value = "HorizontalAlignmentStyle.Center"
        worksheet.Cells(row, 1).Style.HorizontalAlignment = HorizontalAlignmentStyle.Center

        row = row + 2
        worksheet.Cells(row, 0).Value = "Indent"
        worksheet.Cells(row, 1).Value = "five"
        worksheet.Cells(row, 1).Style.HorizontalAlignment = HorizontalAlignmentStyle.Left
        worksheet.Cells(row, 1).Style.Indent = 5

        row = row + 2
        worksheet.Cells(row, 0).Value = "IsTextVertical"
        worksheet.Cells(row, 1).Value = "true"
        worksheet.Rows(row).Height = 60 * 20
        worksheet.Cells(row, 1).Style.IsTextVertical = True

        row = row + 2
        worksheet.Cells(row, 0).Value = "NumberFormat"
        worksheet.Cells(row, 1).Value = 1234
        worksheet.Cells(row, 1).Style.NumberFormat = "#.##0,00 ($Krakozhian Money Units)"

        row = row + 2
        worksheet.Cells(row, 0).Value = "Rotation"
        worksheet.Cells(row, 1).Value = "35 degrees up"
        worksheet.Cells(row, 1).Style.Rotation = 35

        row = row + 2
        worksheet.Cells(row, 0).Value = "ShrinkToFit"
        worksheet.Cells(row, 1).Value = "This property is set to true so this text appears shrunk."
        worksheet.Cells(row, 1).Style.ShrinkToFit = True

        row = row + 2
        worksheet.Cells(row, 0).Value = "VerticalAlignment"
        worksheet.Cells(row, 1).Value = "VerticalAlignmentStyle.Top"
        worksheet.Rows(row).Height = 30 * 20
        worksheet.Cells(row, 1).Style.VerticalAlignment = VerticalAlignmentStyle.Top

        row = row + 2
        worksheet.Cells(row, 0).Value = "WrapText"
        worksheet.Cells(row, 1).Value = "This property is set to true so this text appears broken into multiple lines."
        worksheet.Cells(row, 1).Style.WrapText = True

        workbook.Save("Styles and Formatting.xlsx")
    End Sub
End Module