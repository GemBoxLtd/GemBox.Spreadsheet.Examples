Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Hello World")

        ws.Cells(0, 0).Value = "English:"
        ws.Cells(0, 1).Value = "Hello"

        ws.Cells(1, 0).Value = "Russian:"
        ' Using UNICODE string.
        ws.Cells(1, 1).Value = New String(New Char() {ChrW(&H417), ChrW(&H434), ChrW(&H440), ChrW(&H430), ChrW(&H432), ChrW(&H441), ChrW(&H442), ChrW(&H432), ChrW(&H443), ChrW(&H439), ChrW(&H442), ChrW(&H435)})

        ws.Cells(2, 0).Value = "Chinese:"
        ' Using UNICODE string.
        ws.Cells(2, 1).Value = New String(New Char() {ChrW(&H4F60), ChrW(&H597D)})

        ws.Cells(4, 0).Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC."
        ws.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = True

        ef.Save("Hello World.xlsx")

    End Sub

End Module