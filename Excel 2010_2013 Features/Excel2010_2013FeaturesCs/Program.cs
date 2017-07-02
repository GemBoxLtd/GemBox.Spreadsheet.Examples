using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("Excel 2010.xlsx");

        // Modify all values in column C. Set them to some random value between -10 and 10.
        var readEnumerator = ef.Worksheets[0].Columns["C"].Cells.GetReadEnumerator();

        Random rnd = new Random();
        while (readEnumerator.MoveNext())
        {
            ExcelCell cell = readEnumerator.Current;
            if (cell.ValueType == CellValueType.Int)
                cell.SetValue(rnd.Next(-10, 10));
        }

        ef.Save("Excel 2010_2013 Features.xlsx");
    }
}
