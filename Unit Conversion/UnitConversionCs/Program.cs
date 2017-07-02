using System;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("IllustrationsAndShapes.xlsx");

        StringBuilder sb = new StringBuilder();

        var ws = ef.Worksheets[0];

        sb.AppendFormat("Sheet left margin is: {0} pixels.", Math.Round(LengthUnitConverter.Convert(ws.PrintOptions.LeftMargin, LengthUnit.Inch, LengthUnit.Pixel)));
        sb.AppendLine();

        sb.AppendFormat("Width of column A is: {0} pixels.", Math.Round(LengthUnitConverter.Convert(ws.Columns[0].Width, LengthUnit.ZeroCharacterWidth256thPart, LengthUnit.Pixel)));
        sb.AppendLine();

        sb.AppendFormat("Height of row 1 is: {0} pixels.", Math.Round(LengthUnitConverter.Convert(ws.Rows[0].Height, LengthUnit.Twip, LengthUnit.Pixel)));
        sb.AppendLine();

        var picture = ws.Pictures[1];
        sb.AppendFormat("Image width x height is: {0} centimeters x {1} centimeters.",
            Math.Round(picture.Position.GetWidth(LengthUnit.Centimeter), 2),
            Math.Round(picture.Position.GetHeight(LengthUnit.Centimeter), 2));

        Console.WriteLine(sb.ToString());
    }
}
