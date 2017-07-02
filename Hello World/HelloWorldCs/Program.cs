using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Hello World");

        ws.Cells[0, 0].Value = "English:";
        ws.Cells[0, 1].Value = "Hello";

        ws.Cells[1, 0].Value = "Russian:";
        // Using UNICODE string.
        ws.Cells[1, 1].Value = new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' });

        ws.Cells[2, 0].Value = "Chinese:";
        // Using UNICODE string.
        ws.Cells[2, 1].Value = new string(new char[] { '\u4f60', '\u597d' });

        ws.Cells[4, 0].Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC.";
        ws.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = true;

        ef.Save("Hello World.xlsx");
    }
}
