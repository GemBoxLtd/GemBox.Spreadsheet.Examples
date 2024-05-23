using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Excel workbook from file's path.
        ExcelFile workbook = ExcelFile.Load("CombinedTemplate.xlsx");

        // Set sheets print options.
        foreach (ExcelWorksheet worksheet in workbook.Worksheets)
        {
            ExcelPrintOptions sheetPrintOptions = worksheet.PrintOptions;

            sheetPrintOptions.Portrait = false;
            sheetPrintOptions.HorizontalCentered = true;
            sheetPrintOptions.VerticalCentered = true;

            sheetPrintOptions.PrintHeadings = true;
            sheetPrintOptions.PrintGridlines = true;
        }

        // Create spreadsheet's print options. 
        PrintOptions printOptions = new PrintOptions();
        printOptions.SelectionType = SelectionType.EntireFile;

        // Print Excel workbook to default printer (e.g. 'Microsoft Print to Pdf').
        string printerName = null;
        workbook.Print(printerName, printOptions);
    }
}
