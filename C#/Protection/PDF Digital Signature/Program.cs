using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = "GemBoxExampleExplorer.pfx",
                CertificatePassword = "GemBoxPassword"
            }
        };

        workbook.Save("PDF Digital Signature.pdf", options);
    }
}