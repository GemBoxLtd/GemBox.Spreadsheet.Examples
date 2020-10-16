using System.Linq;
using System.Xml;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var workbook = ExcelFile.Load("HtmlExport.xlsx");

        var worksheet = workbook.Worksheets[0];

        // Set some ExcelPrintOptions properties for HTML export.
        worksheet.PrintOptions.PrintHeadings = true;
        worksheet.PrintOptions.PrintGridlines = true;

        // Specify cell range which should be exported to HTML.
        worksheet.NamedRanges.SetPrintArea(worksheet.Cells.GetSubrange("A1", "J42"));

        var options = new HtmlSaveOptions()
        {
            HtmlType = HtmlType.Html,
            SelectionType = SelectionType.EntireFile
        };

        workbook.Save("HtmlExport.html", options);
    }

    static void Example2()
    {
        var workbook = ExcelFile.Load("HtmlExport.xlsx");

        // Specify exporting of Excel data as an HTML table with embedded images.
        var options = new HtmlSaveOptions()
        {
            EmbedImages = true,
            HtmlType = HtmlType.HtmlTable
        };

        using (var writer = XmlWriter.Create("SingleHtmlExport.html",
            new XmlWriterSettings() { OmitXmlDeclaration = true }))
        {
            writer.WriteStartElement("html");
            writer.WriteStartElement("body");

            // Write Excel sheets to a single HTML file in reverse order.
            foreach (var worksheet in workbook.Worksheets.Reverse())
            {
                writer.WriteElementString("h1", worksheet.Name);

                workbook.Worksheets.ActiveWorksheet = worksheet;
                workbook.Save(writer, options);
            }

            writer.WriteEndDocument();
        }
    }
}