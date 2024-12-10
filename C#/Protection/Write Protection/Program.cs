using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        worksheet.Cells["A1"].Value = "This spreadsheet has been opened in read-only mode.";
        worksheet.Cells["A2"].Value = "Changes cannot be made to the original spreadsheet.";
        worksheet.Cells["A3"].Value = "To save changes a new copy of the spreadsheet must be created.";

        WriteProtection protection = workbook.WriteProtectionSettings;
        protection.SetPassword("pass");

        workbook.Save("XLSX Write Protection.xlsx");
    }
}
