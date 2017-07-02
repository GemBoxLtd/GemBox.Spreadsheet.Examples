Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("ComplexTemplate.xlsx")

        ' In order to achieve the conversion of a loaded Excel file to PDF,
        ' or to some other Excel format,
        ' we just need to save an ExcelFile object to desired output file format.

        ef.Save("Convert.xlsx")

    End Sub

End Module