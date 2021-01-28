using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // In order to convert Excel to PDF, we just need to:
        //   1. Load XLS or XLSX file into ExcelFile object.
        //   2. Save ExcelFile object to PDF file.
        ExcelFile workbook = ExcelFile.Load("ComplexTemplate.xlsx");
        workbook.Save("Convert1.pdf", new PdfSaveOptions() { SelectionType = SelectionType.EntireFile });
    }

    static void Example2()
    {
        // Load Excel file.
        ExcelFile workbook = ExcelFile.Load("ComplexTemplate.xlsx");

        // Get Excel sheet you want to export.
        ExcelWorksheet worksheet = workbook.Worksheets[0];

        // Set targeted sheet as active.
        workbook.Worksheets.ActiveWorksheet = worksheet;

        // Get cell range that you want to export.
        CellRange range = worksheet.Cells.GetSubrange("A5:I14");

        // Set targeted range as print area.
        worksheet.NamedRanges.SetPrintArea(range);

        // Save to PDF file.
        // By default, the SelectionType.ActiveSheet is used.
        workbook.Save("Convert2.pdf");
    }

    static void Example3()
    {
        PdfConformanceLevel conformanceLevel = PdfConformanceLevel.PdfA1a;

        // Load Excel file.
        var workbook = ExcelFile.Load("ComplexTemplate.xlsx");

        // Create PDF save options.
        var options = new PdfSaveOptions()
        {
            ConformanceLevel = conformanceLevel
        };

        // Save to PDF file.
        workbook.Save("Output3.pdf", options);
    }
}