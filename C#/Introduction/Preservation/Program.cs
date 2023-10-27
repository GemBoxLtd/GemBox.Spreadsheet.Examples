using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Excel workbook, preservation feature is enabled by default.
        var workbook = ExcelFile.Load("Preservation.xlsx");
        var worksheet = workbook.Worksheets[0];

        // Modify the worksheet.
        worksheet.Cells["C6"].Value = 8500;
        worksheet.Cells["C7"].Value = 10000;

        // Save Excel worksheet to an output file of the same format together with
        // preserved information (unsupported features) from the input file.
        workbook.Save("PreservedOutput.xlsx");
    }
}
