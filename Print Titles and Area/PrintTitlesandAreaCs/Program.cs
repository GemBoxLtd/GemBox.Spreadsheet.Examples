using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        // Always print 1st row.
        ExcelWorksheet ws1 = ef.Worksheets.Add("Sheet1");
        ws1.NamedRanges.SetPrintTitles(ws1.Rows[0], 1);

        // Set print area (from A1 to I120):
        ws1.NamedRanges.SetPrintArea(ws1.Cells.GetSubrange("A1", "I120"));

        // Always print columns from A to F.
        ExcelWorksheet ws2 = ef.Worksheets.Add("Sheet2");
        ws2.NamedRanges.SetPrintTitles(ws2.Columns[0], 6);

        // Always print columns from A to F and first row.
        ExcelWorksheet ws3 = ef.Worksheets.Add("Sheet3");
        ws3.NamedRanges.SetPrintTitles(ws3.Rows[0], 1, ws3.Columns[0], 6);

        // Fill Sheet1 with some data
        for (int i = 0; i < 9; i++)
            ws1.Cells[0, i].Value = "Column " + ExcelColumnCollection.ColumnIndexToName(i);

        for (int i = 1; i < 120; i++)
            for (int j = 0; j < 9; j++)
                ws1.Cells[i, j].SetValue(i + j);

        ef.Save("Print Titles and Area.xlsx");
    }
}
