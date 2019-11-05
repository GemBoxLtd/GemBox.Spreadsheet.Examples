Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("ComplexTemplate.xlsx")

        ' In order to achieve the conversion of a loaded Excel file to PDF,
        ' or to some other Excel format,
        ' we just need to save an ExcelFile object to desired output file format.
        workbook.Save("Convert.pdf")
    End Sub
End Module