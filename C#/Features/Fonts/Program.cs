using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Set the directory path where the component will look for additional font files.
        // The "MyFonts" targets the subdirectory in the current directory, so besides the installed fonts,
        // the component will be able to use the fonts within the specified directory.
        FontSettings.FontsBaseDirectory = "MyFonts";

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Private Fonts");

        worksheet.Parent.Styles.Normal.Font.Name = "Almonte Snow";
        worksheet.Parent.Styles.Normal.Font.Size = 48 * 20;

        worksheet.Cells[0, 0].Value = "Hello World!";

        workbook.Save("Private Fonts.pdf");
    }
}
