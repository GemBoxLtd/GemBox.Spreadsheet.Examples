Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Referencing")

        ' Referencing cells from sheet using cell names and indexes.
        worksheet.Cells("A1").Value = "Cell A1."
        worksheet.Cells(1, 0).Value = "Cell in 2nd row and 1st column [A2]."

        ' Referencing cells from row using cell names and indexes.
        worksheet.Rows("4").Cells("B").Value = "Cell in row 4 and column B [B4]."
        worksheet.Rows(4).Cells(1).Value = "Cell in 5th row and 2nd column [B5]."

        ' Referencing cells from column using cell names and indexes.
        worksheet.Columns("C").Cells("7").Value = "Cell in column C and row 7 [C7]."
        worksheet.Columns(2).Cells(7).Value = "Cell in 3rd column and 8th row [C8]."

        ' Referencing cell range using A1 notation [G2:N12].
        Dim range = worksheet.Cells.GetSubrange("G2:N12")
        range(0).Value = $"From {range.StartPosition} to {range.EndPosition}"
        range(1, 0).Value = $"From ({range.FirstRowIndex},{range.FirstColumnIndex}) to ({range.LastRowIndex},{range.LastColumnIndex})"
        range.Style.Borders.SetBorders(MultipleBorders.Outside,
            SpreadsheetColor.FromName(ColorName.Red),
            LineStyle.Thick)

        ' Referencing cell range using absolute position [I5:M11].
        range = range.GetSubrangeAbsolute(4, 8, 10, 12)
        range(0).Value = $"From {range.StartPosition} to {range.EndPosition}"
        range(1, 0).Value = $"From ({range.FirstRowIndex},{range.FirstColumnIndex}) to ({range.LastRowIndex},{range.LastColumnIndex})"
        range.Style.Borders.SetBorders(MultipleBorders.Outside,
            SpreadsheetColor.FromName(ColorName.Green),
            LineStyle.Medium)

        ' Referencing cell range using relative position [K8:L10].
        range = range.GetSubrangeRelative(3, 2, 2, 2)
        range(0).Value = $"From {range.StartPosition} to {range.EndPosition}"
        range(1, 0).Value = $"From ({range.FirstRowIndex},{range.FirstColumnIndex}) to ({range.LastRowIndex},{range.LastColumnIndex})"
        range.Style.Borders.SetBorders(MultipleBorders.Outside,
            SpreadsheetColor.FromName(ColorName.Blue),
            LineStyle.Thin)

        workbook.Save("Cell Referencing.xlsx")

    End Sub
End Module