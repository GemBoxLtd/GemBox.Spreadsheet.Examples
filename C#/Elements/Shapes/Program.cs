using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Shapes");

        var roundedRectangle = worksheet.Shapes.Add(ShapeType.RoundedRectangle, "B2", "D4");
        // Radius of the corners is 35% of the rounded rectangle height (since it is smaller than width).
        roundedRectangle.AdjustValues["adj"] = 35000;

        var rightArrow = worksheet.Shapes.Add(ShapeType.RightArrow, "B6", 180, 80, LengthUnit.Point);
        rightArrow.Fill.SetNone();
        rightArrow.Outline.Fill.SetSolid(DrawingColor.FromRgb(250, 30, 20));
        rightArrow.Outline.Width = Length.From(2, LengthUnit.Point);

        var line = worksheet.Shapes.Add(ShapeType.Line, "B12", "B15");
        line.Outline.Width = Length.From(10, LengthUnit.Pixel);

        var shape = worksheet.Shapes.Add(ShapeType.Oval, 100, 100, 200, 150, LengthUnit.Point);
        shape.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.GreenYellow));
        shape.Outline.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.DarkBlue));
        shape.Outline.Width = Length.From(3, LengthUnit.Point);
        // Sending the shape behind the rightArrow.
        shape.SendToBack();

        workbook.Save("Shapes.xlsx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Shapes");

        // Add group.
        var groupShape = worksheet.GroupShapes.Add(100, 50, 200, 250, LengthUnit.Point);
        groupShape.Rotation = 30;

        // Add rounded rectangle.
        var roundedRectangle = groupShape.Shapes.Add(ShapeType.RoundedRectangle, 0, 0, 50, 50, LengthUnit.Point);

        // Add down arrow.
        var downArrowLayout = groupShape.Shapes.Add(ShapeType.DownArrow, 60, 0, 50, 100, LengthUnit.Point);

        // Add picture.
        var picture = groupShape.Pictures.Add("Dices.png", 0, 100, 200, 150, LengthUnit.Point);

        workbook.Save("GroupShapes.xlsx");
    }
}
