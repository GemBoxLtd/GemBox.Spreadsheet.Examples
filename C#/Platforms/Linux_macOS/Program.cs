using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile workbook = new ExcelFile();
        ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet1");
        ExcelCell cell = worksheet.Cells["A1"];

        cell.Value = "Hello World!";

        workbook.Save("HelloWorld.xlsx");
    }
}