using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("Template.xlsx");
        var worksheet = workbook.Worksheets[0];
        var cell = worksheet.Cells["A1"];

        double widthInPoints = cell.Column.GetWidth(LengthUnit.Point);
        double heightInPoints = cell.Row.GetHeight(LengthUnit.Point);

        Console.WriteLine("A1 cell's size in different units:");

        foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
        {
            // The CharacterWidth should not be used with LengthUnitConverter, see:
            // https://www.gemboxsoftware.com/spreadsheet/docs/GemBox.Spreadsheet.LengthUnit.html
            if (unit == LengthUnit.CharacterWidth)
                continue;

            double convertedWidth = LengthUnitConverter.Convert(widthInPoints, LengthUnit.Point, unit);
            double convertedHeight = LengthUnitConverter.Convert(heightInPoints, LengthUnit.Point, unit);
            Console.WriteLine($"{convertedWidth:0.###} x {convertedHeight:0.###} {unit}");
        }
    }
}