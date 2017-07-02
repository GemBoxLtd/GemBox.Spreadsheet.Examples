using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        string inputPassword = "inpass";
        string outputPassword = "outpass";

        var ef = ExcelFile.Load("XlsxEncryption.xlsx", new XlsxLoadOptions() { Password = inputPassword });

        ef.Save("XLSX Encryption.xlsx", new XlsxSaveOptions() { Password = outputPassword });
    }
}
