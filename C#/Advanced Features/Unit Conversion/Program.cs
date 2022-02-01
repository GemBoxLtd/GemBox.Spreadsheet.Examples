using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("Template.xlsx");
        var worksheet = workbook.Worksheets[0];
        var cell = worksheet.Cells["A1"];

        double widthInZeroCharacterWidth256thPart = cell.Column.Width;
        double heightInTwip = cell.Row.Height;

        Console.WriteLine("A1 cell's size in different units:");

        foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
        {
            double convertedWidth = LengthUnitConverter.Convert(widthInZeroCharacterWidth256thPart, LengthUnit.ZeroCharacterWidth256thPart, unit);
            double convertedHeight = LengthUnitConverter.Convert(heightInTwip, LengthUnit.Twip, unit);
            Console.WriteLine($"{convertedWidth:0.###} x {convertedHeight:0.###} {unit}");
        }
    }
}