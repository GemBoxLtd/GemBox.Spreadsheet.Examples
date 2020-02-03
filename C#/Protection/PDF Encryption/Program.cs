using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var password = "pass";
        var ownerPassword = "";

        var options = new PdfSaveOptions()
        {
            DocumentOpenPassword = password,
            PermissionsPassword = ownerPassword,
            Permissions = PdfPermissions.None
        };

        workbook.Save("PDF Encryption.pdf", options);
    }
}