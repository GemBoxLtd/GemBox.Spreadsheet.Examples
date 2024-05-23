using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Security;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var saveOptions = new XlsxSaveOptions();
        saveOptions.DigitalSignatures.Add(new XlsxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword"
        });

        workbook.Save("XLSX Digital Signature.xlsx", saveOptions);
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var signature1 = new XlsxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Created,
            SignerRole = "Developer"
        };
        // Embed intermediate certificate.
        signature1.Certificates.Add(new Certificate("GemBoxECDsa.crt"));

        var signature2 = new XlsxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxRSA4096.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Approved,
            SignerRole = "Manager"
        };
        // Embed intermediate certificate.
        signature2.Certificates.Add(new Certificate("GemBoxRSA.crt"));

        var saveOptions = new XlsxSaveOptions();
        saveOptions.DigitalSignatures.Add(signature1);
        saveOptions.DigitalSignatures.Add(signature2);

        workbook.Save("XLSX Digital Signatures.xlsx", saveOptions);
    }
}
