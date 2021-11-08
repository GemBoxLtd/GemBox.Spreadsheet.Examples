Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim inputPassword = "inpass"
        Dim outputPassword = "outpass"

        Dim workbook = ExcelFile.Load("XlsxEncryption.xlsx",
            New XlsxLoadOptions With {.Password = inputPassword})

        workbook.Save("XLSX Encryption.xlsx",
            New XlsxSaveOptions With {.Password = outputPassword})

    End Sub
End Module