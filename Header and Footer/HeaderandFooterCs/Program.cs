using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Header and Footer");

        var headerFooter = worksheet.HeadersFooters;

        // Show title only on the first page
        headerFooter.FirstPage.Header.CenterSection.Content = "Title on the first page";

        // Show logo
        headerFooter.FirstPage.Header.LeftSection.AppendPicture("Dices.png", 40, 40);
        headerFooter.DefaultPage.Header.LeftSection = headerFooter.FirstPage.Header.LeftSection;

        // "Page number" of "Number of pages"
        headerFooter.FirstPage.Footer.RightSection.Append("Page ").Append(HeaderFooterFieldType.PageNumber).Append(" of ").Append(HeaderFooterFieldType.NumberOfPages);
        headerFooter.DefaultPage.Footer = headerFooter.FirstPage.Footer;

        // Fill Sheet1 with some data
        for (int i = 0; i < 140; i++)
            for (int j = 0; j < 9; j++)
                worksheet.Cells[i, j].Value = i + j;

        workbook.Save("Header and Footer.xlsx");
    }
}