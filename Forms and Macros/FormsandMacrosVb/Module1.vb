Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("FormsAndMacros.xlsm")

        ef.Save("Forms and Macros.xlsm")

    End Sub

End Module