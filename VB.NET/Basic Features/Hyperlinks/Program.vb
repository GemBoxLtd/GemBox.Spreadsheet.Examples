Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Hyperlinks")
        Dim hyperlinkStyle = workbook.Styles(BuiltInCellStyleName.Hyperlink)

        Dim cell = worksheet.Cells("B1")
        cell.Value = "Link to GemBox homepage"
        cell.Style = hyperlinkStyle
        cell.Hyperlink.Location = "https://www.gemboxsoftware.com"
        cell.Hyperlink.IsExternal = True

        cell = worksheet.Cells("B3")
        cell.Value = "Jump"
        cell.Style = hyperlinkStyle
        cell.Hyperlink.ToolTip = "This is tool tip! This hyperlink jumps to E1."
        cell.Hyperlink.Location = worksheet.Name + "!E3"

        worksheet.Cells("E3").Value = "Jump destination"

        cell = worksheet.Cells("B5")
        cell.Formula = "=HYPERLINK(""https://www.gemboxsoftware.com/spreadsheet/examples/excel-cell-hyperlinks/207"", ""Link to Hyperlinks example"")"
        cell.Style = hyperlinkStyle

        workbook.Save("Hyperlinks.xlsx")

    End Sub
End Module