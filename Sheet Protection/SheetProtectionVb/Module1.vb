Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Sheet Protection")

        ws.Cells(0, 2).Value = "Only cells from A1 to A10 are editable."

        For i = 0 To 9 Step 1

            Dim cell = ws.Cells(i, 0)
            cell.SetValue(i)
            cell.Style.Locked = False

        Next

        ws.Protected = True

        ' ProtectionSettings class is supported only for XLSX file format.
        ws.Cells(2, 2).Value = "Inserting columns is allowed (only supported for XLSX file format)."
        Dim protectionSettings = ws.ProtectionSettings
        protectionSettings.AllowInsertingColumns = True

        ws.Cells(3, 2).Value = "Sheet password is 123 (only supported for XLSX file format)."
        protectionSettings.SetPassword("123")

        ef.Save("Sheet Protection.xlsx")

    End Sub

End Module