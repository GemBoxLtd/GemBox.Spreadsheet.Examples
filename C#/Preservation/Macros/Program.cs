using System.IO;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Excel file with preservation feature enabled.
        var loadOptions = new XlsxLoadOptions() { PreserveUnsupportedFeatures = true };
        var workbook = ExcelFile.Load("Macros.xlsm", loadOptions);

        // Save Excel file to output file of same format together with
        // preserved information (unsupported features) from input file.
        string extension = Path.GetExtension("Macros.xlsm");
        workbook.Save($"Preserved Output{extension}");
    }
}
