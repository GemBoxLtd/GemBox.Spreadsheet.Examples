using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("HeadersFooters");

        SheetHeaderFooter sheetHeadersFooters = worksheet.HeadersFooters;

        HeaderFooterPage firstHeaderFooter = sheetHeadersFooters.FirstPage;
        HeaderFooterPage defaultHeaderFooter = sheetHeadersFooters.DefaultPage;

        // Set title text on the center of the first page header.
        firstHeaderFooter.Header.CenterSection
            .Append("Title on the first page",
                new ExcelFont() { Name = "Arial Black", Size = 18 * 20 });

        // Set image on the left of the first and default page headers.
        firstHeaderFooter.Header.LeftSection
            .AppendPicture("Dices.png", 40, 30);
        defaultHeaderFooter.Header.LeftSection = firstHeaderFooter.Header.LeftSection;

        // Set page number on the right of the first and default page footer.
        firstHeaderFooter.Footer.RightSection
            .Append("Page ")
            .Append(HeaderFooterFieldType.PageNumber)
            .Append(" of ")
            .Append(HeaderFooterFieldType.NumberOfPages);
        defaultHeaderFooter.Footer = firstHeaderFooter.Footer;

        worksheet.Cells[0, 0].Value = "First page";
        worksheet.Cells[0, 5].Value = "Second page";
        worksheet.Cells[0, 10].Value = "Third page";

        worksheet.VerticalPageBreaks.Add(5);
        worksheet.VerticalPageBreaks.Add(10);

        workbook.Save("Headers and Footers.xlsx");
    }
}