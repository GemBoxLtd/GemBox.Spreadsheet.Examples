using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");
        var worksheet = workbook.Worksheets.ActiveWorksheet;

        var searchText = "Apollo";
        foreach (var cell in worksheet.Cells.FindAllText(searchText))
            Console.WriteLine($"Text was found in cell '{cell.Name}' (\"{cell.StringValue}\").");
    }
}