using GemBox.Pdf.Forms;
using GemBox.Pdf.Security;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        PAdES_B_B();

        PAdES_B_LTA();
    }

    static void PAdES_B_B()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Create visual representation of digital signature at the beginning of the worksheet.
        var signature = workbook.Worksheets[0].Pictures.Add("GemBoxSignature.png", "B2");

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = "GemBoxECDsa521.pfx",
                CertificatePassword = "GemBoxPassword",
                Signature = signature,
                IsAdvancedElectronicSignature = true
            }
        };

        workbook.Save("PDF Digital Signature.pdf", options);
    }

    static void PAdES_B_LTA()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Create visual representation of digital signature at the beginning of the first worksheet.
        var signature = workbook.Worksheets[0].Pictures.Add("GemBoxSignature.png", "B2");

        // If using Professional version, put your serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Get a digital ID from PKCS#12/PFX file.
        var digitalId = new PdfDigitalId("GemBoxECDsa521.pfx", "GemBoxPassword");

        // Create a PDF signer that will create PAdES B-LTA level signature.
        var signer = new PdfSigner(digitalId);

        // PdfSigner should create CAdES-equivalent signature.
        signer.SignatureFormat = PdfSignatureFormat.CAdES;

        // PdfSigner will embed a timestamp created by freeTSA.org Time Stamp Authority in the signature.
        signer.Timestamper = new PdfTimestamper("https://freetsa.org/tsr");

        // Make sure that all properties specified on PdfSigner are according to PAdES B-LTA level.
        signer.SignatureLevel = PdfSignatureLevel.PAdES_B_LTA;

        // Inject PdfSigner from GemBox.Pdf into
        // PdfDigitalSignatureSaveOptions from GemBox.Spreadsheet.
        var signatureOptions = PdfDigitalSignatureSaveOptions.FromSigner(
            () => signer.SignatureFormat.ToString(),
            () => signer.EstimatedSignatureContentsLength,
            signer.ComputeSignature);

        signatureOptions.Signature = signature;

        var options = new PdfSaveOptions()
        {
            DigitalSignature = signatureOptions
        };

        workbook.Save("PAdES B-LTA.pdf", options);

        using (var pdfDocument = GemBox.Pdf.PdfDocument.Load("PAdES B-LTA.pdf"))
        {
            var signatureField = (PdfSignatureField)pdfDocument.Form.Fields[0];

            // Download validation-related information for the signature and the signature's timestamp and embed it in the PDF file.
            // This will make the signature "LTV enabled".
            pdfDocument.SecurityStore.AddValidationInfo(signatureField.Value);

            // Add an invisible signature field to the PDF document that will hold the document timestamp.
            var timestampField = pdfDocument.Form.Fields.AddSignature();

            // Initiate timestamping of a PDF file with the specified timestamper.
            timestampField.Timestamp(signer.Timestamper);

            // Save any changes done to the PDF file that were done since the last time Save was called and
            // finish timestamping of a PDF file.
            pdfDocument.Save();
        }
    }
}