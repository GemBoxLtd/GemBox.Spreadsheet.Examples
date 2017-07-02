using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Define columns width (for input file format)
        FixedWidthLoadOptions loadOptions = new FixedWidthLoadOptions(
            new FixedWidthColumn(8),
            new FixedWidthColumn(8),
            new FixedWidthColumn(8));

        // Load file
        ExcelFile ef = ExcelFile.Load("FixedColumnsWidthText.prn", loadOptions);

        // Modify file
        ef.Worksheets.ActiveWorksheet.GetUsedCellRange(true).Sort(false).By(1).Apply();

        // Define columns width (for output file format)
        FixedWidthSaveOptions saveOptions = new FixedWidthSaveOptions(
            new FixedWidthColumn(8),
            new FixedWidthColumn(8),
            new FixedWidthColumn(8));

        ef.Save("FixedColumnsWidthText.prn", saveOptions);
    }
}
