using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var inputPassword = "inpass";
        var outputPassword = "outpass";

        var workbook = ExcelFile.Load("XlsxEncryption.xlsx", new XlsxLoadOptions() { Password = inputPassword });

        workbook.Save("XLSX Encryption.xlsx", new XlsxSaveOptions() { Password = outputPassword });
    }
}