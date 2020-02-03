Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("FormsAndMacros.xlsm")

        workbook.Save("Forms and Macros.xlsm")
    End Sub
End Module