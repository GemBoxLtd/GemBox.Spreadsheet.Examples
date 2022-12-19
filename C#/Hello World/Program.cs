using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Hello World");

        worksheet.Cells[0, 0].Value = "English:";
        worksheet.Cells[0, 1].Value = "Hello";

        worksheet.Cells[1, 0].Value = "Russian:";
        // Using UNICODE string.
        worksheet.Cells[1, 1].Value = new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' });

        worksheet.Cells[2, 0].Value = "Chinese:";
        // Using UNICODE string.
        worksheet.Cells[2, 1].Value = new string(new char[] { '\u4f60', '\u597d' });

        worksheet.Cells[4, 0].Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC.";
        worksheet.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = true;

        workbook.Save("Hello World.xlsx");
    }
}