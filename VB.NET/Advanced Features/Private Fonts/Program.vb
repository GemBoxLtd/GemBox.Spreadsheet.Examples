Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Private Fonts")

        ' Current directory contains a font file.
        FontSettings.FontsBaseDirectory = "."

        worksheet.Parent.Styles.Normal.Font.Name = "Almonte Snow"
        worksheet.Cells(0, 0).Value = "Hello World!"

        workbook.Save("Private Fonts.pdf")
    End Sub
End Module