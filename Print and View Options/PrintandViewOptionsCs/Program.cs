using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Print and View Options");

        ws.Cells["M1"].Value = "This worksheet shows how to set various print related and view related options.";
        ws.Cells["M2"].Value = "To see results of print options, go to Print and Page Setup dialogs in MS Excel.";
        ws.Cells["M3"].Value = "Notice that print and view options are worksheet based, not workbook based.";

        // Print options:
        var printOptions = ws.PrintOptions;
        printOptions.PrintGridlines = true;
        printOptions.PrintHeadings = true;
        printOptions.Portrait = false;
        printOptions.PaperType = PaperType.A3;
        printOptions.NumberOfCopies = 5;

        // View options:
        ws.ViewOptions.FirstVisibleColumn = 3;
        ws.ViewOptions.ShowColumnsFromRightToLeft = true;
        ws.ViewOptions.Zoom = 123;

        // Set print area
        ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("E1", "U7"));

        ef.Save("Print and View Options.xlsx");
    }
}
