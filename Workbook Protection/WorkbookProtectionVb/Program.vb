Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Workbook Protection")

        ' ProtectionSettings class is supported only for XLSX file format.
        Dim protectionSettings = workbook.ProtectionSettings
        protectionSettings.ProtectStructure = True

        worksheet.Cells(0, 0).Value = "Workbook password is 123 (only supported for XLSX file format)."
        protectionSettings.SetPassword("123")

        workbook.Save("Workbook Protection.xlsx")
    End Sub
End Module