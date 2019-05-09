using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("FormsAndMacros.xlsm");

        workbook.Save("Forms and Macros.xlsm");
    }
}
