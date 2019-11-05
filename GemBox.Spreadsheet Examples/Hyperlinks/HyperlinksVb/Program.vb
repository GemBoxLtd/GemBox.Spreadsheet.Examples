Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Hyperlinks")

        worksheet.Cells("A1").Value = "Hyperlink examples:"

        With worksheet.Cells.Item("B3")
            .Value = "GemboxSoftware"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
            .Hyperlink.Location = "https://www.gemboxsoftware.com"
            .Hyperlink.IsExternal = True
        End With

        With worksheet.Cells.Item("B5")
            .Value = "Jump"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
            .Hyperlink.ToolTip = "This is tool tip! This hyperlink jumps to E1!"
            .Hyperlink.Location = worksheet.Name & "!E1"
        End With

        worksheet.Cells("E1").Value = "Destination"

        With worksheet.Cells.Item("B8")
            .Formula = "=HYPERLINK(""https://www.gemboxsoftware.com/spreadsheet/examples/excel-cell-hyperlinks/207"", ""Example of HYPERLINK formula"")"
            .Style.Font.UnderlineStyle = UnderlineStyle.Single
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue)
        End With

        workbook.Save("Hyperlinks.xlsx")
    End Sub
End Module