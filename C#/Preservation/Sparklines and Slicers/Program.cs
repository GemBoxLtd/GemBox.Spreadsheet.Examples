using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // Load Excel file with preservation feature enabled.
        var loadOptions = new XlsxLoadOptions() { PreserveUnsupportedFeatures = true };
        var workbook = ExcelFile.Load("SparklinesAndSlicers.xlsx", loadOptions);

        // Modify all values in column C, set them to some random value.
        var readEnumerator = workbook.Worksheets[0].Columns["C"].Cells.GetReadEnumerator();
        var random = new Random();
        while (readEnumerator.MoveNext())
        {
            var cell = readEnumerator.Current;
            if (cell.ValueType == CellValueType.Int)
                cell.SetValue(random.Next(-10, 10));
        }

        workbook.Save("Preserved Output.xlsx");
    }
}