Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim password = "pass"
        Dim ownerPassword = ""

        Dim options = New PdfSaveOptions() With
        {
            .DocumentOpenPassword = password,
            .PermissionsPassword = ownerPassword,
            .Permissions = PdfPermissions.None
        }

        workbook.Save("PDF Encryption.pdf", options)
    End Sub
End Module
