using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Private Fonts");

        string pathToResources = "Resources";

        FontSettings.FontsBaseDirectory = pathToResources;

        ws.Parent.Styles.Normal.Font.Name = "Almonte Snow";
        ws.Cells[0, 0].Value = "Hello World!";

        ef.Save("Private Fonts.pdf");
    }
}
