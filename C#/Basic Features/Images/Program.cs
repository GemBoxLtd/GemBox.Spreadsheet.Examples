using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
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
}