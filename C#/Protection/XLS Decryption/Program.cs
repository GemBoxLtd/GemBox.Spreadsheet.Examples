using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var inputPassword = "inpass";

        var workbook = ExcelFile.Load("XlsDecryption.xls",
            new XlsLoadOptions() { Password = inputPassword });

        workbook.Save("Decrypted File.xlsx");
    }
}