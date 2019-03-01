Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("HtmlExport.xlsx")

        Dim worksheet = workbook.Worksheets(0)

        ' Some of the properties from ExcelPrintOptions class are supported in HTML export.
        worksheet.PrintOptions.PrintHeadings = True
        worksheet.PrintOptions.PrintGridlines = True

        ' Print area can be used to specify custom cell range which should be exported to HTML.
        worksheet.NamedRanges.SetPrintArea(worksheet.Cells.GetSubrange("A1", "J42"))

        Dim options = New HtmlSaveOptions() With
        {
            .HtmlType = HtmlType.Html,
            .SelectionType = SelectionType.EntireFile
        }

        workbook.Save("HtmlExport.html", options)
    End Sub
End Module