using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("HtmlExport.xlsx");

        var ws = ef.Worksheets[0];

        // Some of the properties from ExcelPrintOptions class are supported in HTML export.
        ws.PrintOptions.PrintHeadings = true;
        ws.PrintOptions.PrintGridlines = true;

        // Print area can be used to specify custom cell range which should be exported to HTML.
        ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("A1", "I42"));

        HtmlSaveOptions options = new HtmlSaveOptions()
        {
            HtmlType = HtmlType.Html,
            SelectionType = SelectionType.EntireFile
        };

        ef.Save("HtmlExport.html", options);
    }
}
