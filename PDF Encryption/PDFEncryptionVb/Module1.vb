Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim password As String = "pass"
        Dim ownerPassword As String = ""

        Dim options = New PdfSaveOptions() With
        {
            .DocumentOpenPassword = password,
            .PermissionsPassword = ownerPassword,
            .Permissions = PdfPermissions.None
        }

        ef.Save("PDF Encryption.pdf", options)

    End Sub

End Module