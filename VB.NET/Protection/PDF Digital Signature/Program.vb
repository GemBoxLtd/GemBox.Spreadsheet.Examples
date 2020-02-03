Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim options = New PdfSaveOptions()
        Dim digitalSignature = options.DigitalSignature

        digitalSignature.CertificatePath = "GemBoxExampleExplorer.pfx"
        digitalSignature.CertificatePassword = "GemBoxPassword"

        workbook.Save("PDF Digital Signature.pdf", options)
    End Sub
End Module