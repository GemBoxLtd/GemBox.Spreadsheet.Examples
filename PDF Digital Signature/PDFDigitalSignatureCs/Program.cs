using System.IO;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        string pathToResources = "Resources";

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = Path.Combine(pathToResources, "GemBoxSampleExplorer.pfx"),
                CertificatePassword = "GemBoxPassword"
            }
        };

        ef.Save("PDF Digital Signature.pdf", options);
    }
}
