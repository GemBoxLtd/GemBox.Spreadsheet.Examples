Imports System.IO
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim pathToResources As String = "Resources"

        Dim options = New PdfSaveOptions()
        Dim digitalSignature = options.DigitalSignature

        digitalSignature.CertificatePath = Path.Combine(pathToResources, "GemBoxSampleExplorer.pfx")
        digitalSignature.CertificatePassword = "GemBoxPassword"

        ef.Save("PDF Digital Signature.pdf", options)

    End Sub

End Module