using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Sheet Protection");

        ws.Cells[0, 2].Value = "Only cells from A1 to A10 are editable.";

        for (int i = 0; i < 10; i++)
        {
            var cell = ws.Cells[i, 0];
            cell.SetValue(i);
            cell.Style.Locked = false;
        }

        ws.Protected = true;

        // ProtectionSettings class is supported only for XLSX file format.
        ws.Cells[2, 2].Value = "Inserting columns is allowed (only supported for XLSX file format).";
        var protectionSettings = ws.ProtectionSettings;
        protectionSettings.AllowInsertingColumns = true;

        ws.Cells[3, 2].Value = "Sheet password is 123 (only supported for XLSX file format).";
        protectionSettings.SetPassword("123");

        ef.Save("Sheet Protection.xlsx");
    }
}
