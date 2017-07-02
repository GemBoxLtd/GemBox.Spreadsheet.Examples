using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        var ws = ef.Worksheets[0];

        int columnCount = ws.CalculateMaxUsedColumns();
        for (int i = 0; i < columnCount; i++)
            ws.Columns[i].AutoFit(1, ws.Rows[1], ws.Rows[ws.Rows.Count - 1]);

        ef.Save("Row_Column AutoFit.pdf");
    }
}
