Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Security

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim saveOptions As New XlsxSaveOptions()
        saveOptions.DigitalSignatures.Add(New XlsxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword"
        })

        workbook.Save("XLSX Digital Signature.xlsx", saveOptions)
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim signature1 As New XlsxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Created,
            .SignerRole = "Developer"
        }
        ' Embed intermediate certificate.
        signature1.Certificates.Add(New Certificate("GemBoxECDsa.crt"))

        Dim signature2 As New XlsxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxRSA4096.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Approved,
            .SignerRole = "Manager"
        }
        ' Embed intermediate certificate.
        signature2.Certificates.Add(New Certificate("GemBoxRSA.crt"))

        Dim saveOptions As New XlsxSaveOptions()
        saveOptions.DigitalSignatures.Add(signature1)
        saveOptions.DigitalSignatures.Add(signature2)

        workbook.Save("XLSX Digital Signatures.xlsx", saveOptions)
    End Sub
End Module
