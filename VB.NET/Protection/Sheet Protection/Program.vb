Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Sheet Protection")

        worksheet.Cells(0, 2).Value = "Only cells from A1 to A10 are editable."

        For i = 0 To 9 Step 1
            Dim cell = worksheet.Cells(i, 0)
            cell.SetValue(i)
            cell.Style.Locked = False
        Next

        worksheet.Protected = True

        Dim protectionSettings = worksheet.ProtectionSettings
        worksheet.Cells(3, 2).Value = "Sheet password is 123 (only supported for XLSX and XLS file format)."
        protectionSettings.SetPassword("123")

        worksheet.Cells(2, 2).Value = "Inserting columns is allowed (only supported for XLSX file format)."
        protectionSettings.AllowInsertingColumns = True

        workbook.Save("Sheet Protection.xlsx")

    End Sub
End Module
