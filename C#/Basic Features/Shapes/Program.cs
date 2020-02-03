using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Shapes");

        var shape = worksheet.Shapes.Add(ShapeType.Oval, 100, 100, 200, 150, LengthUnit.Point);
        shape.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.GreenYellow));
        shape.Outline.Fill.SetSolid(DrawingColor.FromName(DrawingColorName.DarkBlue));
        shape.Outline.Width = Length.From(3, LengthUnit.Point);

        var roundedRectangle = worksheet.Shapes.Add(ShapeType.RoundedRectangle, "B2", "D4");
        // Radius of the corners is 35% of the rounded rectangle height (since it is smaller than width).
        roundedRectangle.AdjustValues["adj"] = 35000;

        var rightArrow = worksheet.Shapes.Add(ShapeType.RightArrow, "B6", 100, 40, LengthUnit.Point);
        rightArrow.Fill.SetNone();
        rightArrow.Outline.Fill.SetSolid(DrawingColor.FromRgb(250, 30, 20));
        rightArrow.Outline.Width = Length.From(2, LengthUnit.Point);

        var line = worksheet.Shapes.Add(ShapeType.Line, "B12", "B15");
        line.Outline.Width = Length.From(10, LengthUnit.Pixel);

        worksheet.PrintOptions.PrintGridlines = true;
        worksheet.PrintOptions.PrintHeadings = true;

        workbook.Save("Shapes.xlsx");
    }
}