using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Private Fonts");

        // Current directory contains a font file.
        FontSettings.FontsBaseDirectory = ".";

        worksheet.Parent.Styles.Normal.Font.Name = "Almonte Snow";
        worksheet.Parent.Styles.Normal.Font.Size = 48 * 20;

        worksheet.Cells[0, 0].Value = "Hello World!";

        workbook.Save("Private Fonts.pdf");
    }
}