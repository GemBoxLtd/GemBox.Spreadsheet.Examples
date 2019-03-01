Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("IllustrationsAndShapes.xlsx")

        workbook.Save("Illustrations and Shapes.xlsx")
    End Sub
End Module