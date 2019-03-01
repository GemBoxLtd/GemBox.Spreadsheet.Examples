using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("HtmlExport.xlsx");

        var worksheet = workbook.Worksheets[0];

        // Some of the properties from ExcelPrintOptions class are supported in HTML export.
        worksheet.PrintOptions.PrintHeadings = true;
        worksheet.PrintOptions.PrintGridlines = true;

        // Print area can be used to specify custom cell range which should be exported to HTML.
        worksheet.NamedRanges.SetPrintArea(worksheet.Cells.GetSubrange("A1", "J42"));

        var options = new HtmlSaveOptions()
        {
            HtmlType = HtmlType.Html,
            SelectionType = SelectionType.EntireFile
        };

        workbook.Save("HtmlExport.html", options);
    }
}