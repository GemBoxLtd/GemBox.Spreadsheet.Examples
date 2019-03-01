using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("IllustrationsAndShapes.xlsx");

        workbook.Save("Illustrations and Shapes.xlsx");
    }
}
