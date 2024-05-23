using GemBox.Spreadsheet;
using System.Linq;
using System.Xml;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

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
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

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
                if (worksheet.Visibility != SheetVisibility.Visible)
                    continue;

                writer.WriteElementString("h1", worksheet.Name);
                workbook.Worksheets.ActiveWorksheet = worksheet;
                workbook.Save(writer, options);
            }

            writer.WriteEndDocument();
        }
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Load input HTML file.
        var workbook = ExcelFile.Load("HtmlImport.html");

        // Save output XLSX file.
        workbook.Save("HtmlImport.xlsx");
    }
}
