using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("ComplexTemplate.xlsx");

        // In order to achieve the conversion of a loaded Excel file to PDF,
        // or to some other Excel format,
        // we just need to save an ExcelFile object to desired output file format.

        ef.Save("Convert.xlsx");
    }
}
