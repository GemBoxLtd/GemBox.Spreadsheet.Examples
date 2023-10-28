Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("HeadersFooters")

        Dim sheetHeadersFooters As SheetHeaderFooter = worksheet.HeadersFooters

        Dim firstHeaderFooter As HeaderFooterPage = sheetHeadersFooters.FirstPage
        Dim defaultHeaderFooter As HeaderFooterPage = sheetHeadersFooters.DefaultPage

        ' Set title text on the center of the first page header.
        firstHeaderFooter.Header.CenterSection _
            .Append("Title on the first page",
                New ExcelFont() With {.Name = "Arial Black", .Size = 18 * 20})

        ' Set image on the left of the first and default page headers.
        firstHeaderFooter.Header.LeftSection _
            .AppendPicture("Dices.png", 40, 30)
        defaultHeaderFooter.Header.LeftSection = firstHeaderFooter.Header.LeftSection

        ' Set page number on the right of the first and default page footer.
        firstHeaderFooter.Footer.RightSection _
            .Append("Page ") _
            .Append(HeaderFooterFieldType.PageNumber) _
            .Append(" of ") _
            .Append(HeaderFooterFieldType.NumberOfPages)
        defaultHeaderFooter.Footer = firstHeaderFooter.Footer

        worksheet.Cells(0, 0).Value = "First page"
        worksheet.Cells(0, 5).Value = "Second page"
        worksheet.Cells(0, 10).Value = "Third page"

        worksheet.VerticalPageBreaks.Add(5)
        worksheet.VerticalPageBreaks.Add(10)

        workbook.Save("Headers and Footers.xlsx")

    End Sub
End Module