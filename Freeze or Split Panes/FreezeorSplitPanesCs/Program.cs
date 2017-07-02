using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        
        ExcelFile ef = new ExcelFile();

        // Frozen Rows (first 2 rows are frozen)
        ExcelWorksheet ws1 = ef.Worksheets.Add("Frozen rows");
        ws1.Panes = new WorksheetPanes(PanesState.Frozen, 0, 2, "A3", PanePosition.BottomLeft);

        // Frozen Columns (first column is frozen)
        ExcelWorksheet ws2 = ef.Worksheets.Add("Frozen columns");
        ws2.Panes = new WorksheetPanes(PanesState.Frozen, 1, 0, "B1", PanePosition.TopRight);

        // Frozen Rows and Columns (first 2 rows and first 3 columns are frozen)
        ExcelWorksheet ws3 = ef.Worksheets.Add("Frozen rows and columns");
        ws3.Panes = new WorksheetPanes(PanesState.Frozen, 3, 2, "E5", PanePosition.BottomRight);

        // Split pane
        ExcelWorksheet ws4 = ef.Worksheets.Add("Split pane");
        ws4.Panes = new WorksheetPanes(PanesState.Split, 2310, 1500, "D7", PanePosition.BottomRight);

        ef.Save("Freeze or Split Panes.xlsx");
    }
}
