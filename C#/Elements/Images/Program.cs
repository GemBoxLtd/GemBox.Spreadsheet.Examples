using GemBox.Spreadsheet;
using System;

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
        var worksheet = workbook.Worksheets.Add("Images");

        // Add small BMP image with specified rectangle position.
        worksheet.Pictures.Add("SmallImage.bmp", 50, 50, 48, 48, LengthUnit.Pixel);

        // Add large JPG image with specified top-left cell.
        worksheet.Pictures.Add("FragonardReader.jpg", "B9");

        // Add PNG image with specified top-left and bottom-right cells.
        worksheet.Pictures.Add("Dices.png", "J16", "K20");

        // Add GIF image using anchors.
        var picture = worksheet.Pictures.Add("Zahnrad.gif",
            new AnchorCell(worksheet.Columns[9], worksheet.Rows[21], 100000, 100000),
            new AnchorCell(worksheet.Columns[10], worksheet.Rows[23], 50000, 50000));

        // Set picture's position mode.
        picture.Position.Mode = PositioningMode.Move;

        // Add SVG image with specified top-left cell and size.
        picture = worksheet.Pictures.Add("Graphics1.svg", "J9", 250, 100, LengthUnit.Pixel);

        // Set picture's metadata.
        picture.Metadata.Name = "SVG Image";

        workbook.Save("Images.xlsx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Smileys");

        // Create a sheet with specified columns width and rows height.
        for (int i = 0; i < 6; i++)
        {
            worksheet.Columns[i].SetWidth(10 * (i + 1), LengthUnit.Point);
            worksheet.Rows[i].SetHeight(10 * (i + 1), LengthUnit.Point);
        }

        // Add images that fit inside a single cell.
        foreach (var cell in worksheet.Cells.GetSubrange("A1:F6"))
        {
            var picture = worksheet.Pictures.Add("SmilingFace.png", cell.Name);
            var position = picture.Position;

            double maxWidth = cell.Column.GetWidth(LengthUnit.Point);
            double maxHeight = cell.Row.GetHeight(LengthUnit.Point);

            var ratioX = maxWidth / position.Width;
            var ratioY = maxHeight / position.Height;
            var ratio = Math.Min(ratioX, ratioY);

            if (ratio < 1)
            {
                position.Width *= ratio;
                position.Height *= ratio;
            }
        }

        workbook.Save("CellsImages.xlsx");
    }
}
