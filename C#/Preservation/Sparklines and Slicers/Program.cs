using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("ExcelSpecific.xlsx");

        // Modify all values in column C. Set them to some random value between -10 and 10.
        var readEnumerator = workbook.Worksheets[0].Columns["C"].Cells.GetReadEnumerator();

        var rnd = new Random();
        while (readEnumerator.MoveNext())
        {
            var cell = readEnumerator.Current;
            if (cell.ValueType == CellValueType.Int)
                cell.SetValue(rnd.Next(-10, 10));
        }

        workbook.Save("Excel Specific Features.xlsx");
    }
}