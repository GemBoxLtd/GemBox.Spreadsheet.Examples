Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim inputPassword = "inpass"

        Dim workbook = ExcelFile.Load("XlsDecryption.xls",
            New XlsLoadOptions With {.Password = inputPassword})

        workbook.Save("Decrypted File.xls")

    End Sub
End Module
