Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Header and Footer")

        Dim headerFooter = worksheet.HeadersFooters

        ' Show title only on the first page
        headerFooter.FirstPage.Header.CenterSection.Content = "Title on the first page"

        ' Show logo
        headerFooter.FirstPage.Header.LeftSection.AppendPicture("Dices.png", 40, 40)
        headerFooter.DefaultPage.Header.LeftSection = headerFooter.FirstPage.Header.LeftSection

        ' "Page number" of "Number of pages"
        headerFooter.FirstPage.Footer.RightSection.Append("Page ").Append(HeaderFooterFieldType.PageNumber).Append(" of ").Append(HeaderFooterFieldType.NumberOfPages)
        headerFooter.DefaultPage.Footer = headerFooter.FirstPage.Footer

        ' Fill Sheet1 with some data
        For i As Integer = 0 To 139
            For j As Integer = 0 To 8
                worksheet.Cells(i, j).Value = i + j
            Next
        Next

        workbook.Save("Header and Footer.xlsx")
    End Sub
End Module