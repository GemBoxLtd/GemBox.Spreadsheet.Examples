using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Define columns width (for input file format).
        var loadOptions = new FixedWidthLoadOptions(
            new FixedWidthColumn(8),
            new FixedWidthColumn(8),
            new FixedWidthColumn(8));

        // Load file.
        var workbook = ExcelFile.Load("FixedColumnsWidthText.prn", loadOptions);

        // Modify file.
        workbook.Worksheets.ActiveWorksheet.GetUsedCellRange(true).Sort(false).By(1).Apply();

        // Define columns width (for output file format).
        var saveOptions = new FixedWidthSaveOptions(
            new FixedWidthColumn(8),
            new FixedWidthColumn(8),
            new FixedWidthColumn(8));

        workbook.Save("Fixed Columns Width Text.prn", saveOptions);
    }
}