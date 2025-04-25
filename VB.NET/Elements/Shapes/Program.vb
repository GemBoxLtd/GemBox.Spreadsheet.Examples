Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Drawing

Module Program

    Sub Main()
        Example1()
        Example2()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Shapes")

        Dim roundedRectangle = worksheet.Shapes.Add(ShapeType.RoundedRectangle, "B2", "D4")
        ' Radius of the corners is 35% of the rounded rectangle height (since it is smaller than width).
        roundedRectangle.AdjustValues.Add("adj", 35000)

        Dim rightArrow = worksheet.Shapes.Add(ShapeType.RightArrow, "B6", 100, 40, LengthUnit.Point)
        rightArrow.Fill.SetNone()
        rightArrow.Outline.Fill.SetSolid(DrawingColor.FromRgb(250, 30, 20))
        rightArrow.Outline.Width = Length.From(2, LengthUnit.Point)

        Dim line = worksheet.Shapes.Add(ShapeType.Line, "B12", "B15")
        line.Outline.Width = Length.From(10, LengthUnit.Pixel)

        Dim shape = worksheet.Shapes.Add(ShapeType.Oval, 100, 100, 200, 150, LengthUnit.Point)
        shape.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.GreenYellow))
        shape.Outline.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.DarkBlue))
        shape.Outline.Width = Length.From(3, LengthUnit.Point)
        ' Sending the shape behind the rightArrow.
        shape.SendToBack()

        workbook.Save("Shapes.xlsx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Shapes")

        ' Add group.
        Dim groupShape = worksheet.GroupShapes.Add(100, 50, 200, 250, LengthUnit.Point)
        groupShape.Rotation = 30

        ' Add rounded rectangle.
        Dim roundedRectangle = groupShape.Shapes.Add(ShapeType.RoundedRectangle, 0, 0, 50, 50, LengthUnit.Point)

        ' Add down arrow.
        Dim downArrowLayout = groupShape.Shapes.Add(ShapeType.DownArrow, 60, 0, 50, 100, LengthUnit.Point)

        ' Add picture.
        Dim picture = groupShape.Pictures.Add("Dices.png", 0, 100, 200, 150, LengthUnit.Point)

        workbook.Save("GroupShapes.xlsx")
    End Sub

End Module
