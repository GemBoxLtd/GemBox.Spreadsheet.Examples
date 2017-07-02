Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Hyperlinks")

        ws.Cells("A1").Value = "Hyperlink examples:"

        With ws.Cells.Item("B3")
            .Value = "GemboxSoftware"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
            .Hyperlink.Location = "https://www.gemboxsoftware.com"
            .Hyperlink.IsExternal = True
        End With

        With ws.Cells.Item("B5")
            .Value = "Jump"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
            .Hyperlink.ToolTip = "This is tool tip! This hyperlink jumps to E1!"
            .Hyperlink.Location = (ws.Name & "!E1")
        End With

        ws.Cells("E1").Value = "Destination"

        With ws.Cells.Item("B8")
            .Formula = "=HYPERLINK(""https://www.gemboxsoftware.com/spreadsheet/examples/excel-cell-hyperlinks/207"", ""Example of HYPERLINK formula"")"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
        End With

        ef.Save("Hyperlinks.xls")

    End Sub

End Module