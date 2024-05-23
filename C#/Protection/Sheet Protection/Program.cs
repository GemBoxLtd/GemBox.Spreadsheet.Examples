using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sheet Protection");

        worksheet.Cells[0, 2].Value = "Only cells from A1 to A10 are editable.";

        for (int i = 0; i < 10; i++)
        {
            var cell = worksheet.Cells[i, 0];
            cell.SetValue(i);
            cell.Style.Locked = false;
        }

        worksheet.Protected = true;

        worksheet.Cells[2, 2].Value = "Inserting columns is allowed (only supported for XLSX file format).";
        var protectionSettings = worksheet.ProtectionSettings;
        protectionSettings.AllowInsertingColumns = true;

        worksheet.Cells[3, 2].Value = "Sheet password is 123 (only supported for XLSX file format).";
        protectionSettings.SetPassword("123");

        workbook.Save("Sheet Protection.xlsx");
    }
}
