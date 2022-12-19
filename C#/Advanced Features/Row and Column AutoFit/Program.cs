using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var worksheet = workbook.Worksheets[0];

        int columnCount = worksheet.CalculateMaxUsedColumns();
        for (int i = 0; i < columnCount; i++)
            worksheet.Columns[i].AutoFit(1, worksheet.Rows[1], worksheet.Rows[worksheet.Rows.Count - 1]);

        workbook.Save("Row_Column AutoFit.xlsx");
    }
}