using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");
        var worksheet = workbook.Worksheets[0];

        var searchText = "Apollo";
        var range = worksheet.Columns[0].Cells;

        while (range.FindText(searchText, out int row, out int column))
        {
            var cell = worksheet.Cells[row, column];
            Console.WriteLine($"Text was found in cell '{cell.Name}' (\"{cell.StringValue}\").");

            range = range.GetSubrangeAbsolute(row + 1, 0, worksheet.Rows.Count, 0);
        }
    }
}