Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Private Fonts")

        Dim pathToResources As String = "Resources"

        FontSettings.FontsBaseDirectory = pathToResources

        ws.Parent.Styles.Normal.Font.Name = "Almonte Snow"
        ws.Cells(0, 0).Value = "Hello World!"

        ef.Save("Private Fonts.pdf")

    End Sub

End Module