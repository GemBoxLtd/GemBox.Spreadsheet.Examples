Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Drawing

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Shapes")

        Dim shape = worksheet.Shapes.Add(ShapeType.Oval, 100, 100, 200, 150, LengthUnit.Point)
        shape.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.GreenYellow))
        shape.Outline.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.DarkBlue))
        shape.Outline.Width = Length.From(3, LengthUnit.Point)

        Dim roundedRectangle = worksheet.Shapes.Add(ShapeType.RoundedRectangle, "B2", "D4")
        ' Radius of the corners is 35% of the rounded rectangle height (since it is smaller than width).
        roundedRectangle.AdjustValues.Add("adj", 35000)

        Dim rightArrow = worksheet.Shapes.Add(ShapeType.RightArrow, "B6", 100, 40, LengthUnit.Point)
        rightArrow.Fill.SetNone()
        rightArrow.Outline.Fill.SetSolid(DrawingColor.FromRgb(250, 30, 20))
        rightArrow.Outline.Width = Length.From(2, LengthUnit.Point)

        Dim line = worksheet.Shapes.Add(ShapeType.Line, "B12", "B15")
        line.Outline.Width = Length.From(10, LengthUnit.Pixel)

        workbook.Save("Shapes.xlsx")
    End Sub
End Module