Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim inputPassword As String = "inpass"
        Dim outputPassword As String = "outpass"

        Dim ef = ExcelFile.Load("XlsxEncryption.xlsx", New XlsxLoadOptions With {.Password = inputPassword})

        ef.Save("XLSX Encryption.xlsx", New XlsxSaveOptions With {.Password = outputPassword})

    End Sub

End Module