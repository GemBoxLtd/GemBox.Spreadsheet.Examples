Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef = ExcelFile.Load("HtmlExport.xlsx")

        Dim ws = ef.Worksheets(0)

        ' Some of the properties from ExcelPrintOptions class are supported in HTML export.
        ws.PrintOptions.PrintHeadings = True
        ws.PrintOptions.PrintGridlines = True

        ' Print area can be used to specify custom cell range which should be exported to HTML.
        ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "I42"))

        Dim options = New HtmlSaveOptions() With
        {
            .HtmlType = HtmlType.Html,
            .SelectionType = SelectionType.EntireFile
        }

        ef.Save("HtmlExport.html", options)

    End Sub

End Module