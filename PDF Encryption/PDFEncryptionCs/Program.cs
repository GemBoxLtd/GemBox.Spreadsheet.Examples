using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        string password = "pass";
        string ownerPassword = "";

        var options = new PdfSaveOptions()
        {
            DocumentOpenPassword = password,
            PermissionsPassword = ownerPassword,
            Permissions = PdfPermissions.None
        };

        ef.Save("PDF Encryption.pdf", options);
    }
}
