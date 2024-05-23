using GemBox.Spreadsheet;
using System;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");
        var worksheet = workbook.Worksheets.ActiveWorksheet;

        // Find first cell with specific text.
        var searchText = "Ranger";
        if (worksheet.Cells.FindText(searchText, out ExcelCell foundCell))
        {
            Console.WriteLine($"First cell with '{searchText}' text:");
            Console.WriteLine($"Name: {foundCell.Name} | Value: \"{foundCell.StringValue}\"");
            Console.WriteLine();
        }

        // Find all cells with specific text.
        searchText = "Apollo";
        Console.WriteLine($"All cells with '{searchText}' text:");

        foreach (var cell in worksheet.Cells.FindAllText(searchText))
            Console.WriteLine($"Name: {cell.Name} | Value: \"{cell.StringValue}\"");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");
        var worksheet = workbook.Worksheets.ActiveWorksheet;

        // Replace specific text in first cell in which it occurs.
        string searchText = "Ranger";
        if (worksheet.Cells.FindText(searchText, out ExcelCell foundCell))
            foundCell.ReplaceText(searchText, "REPLACED FIRST");

        // Replace specific text in all cells in which it occurs.
        worksheet.Cells.ReplaceText("Apollo", "REPLACED ALL");

        // Replace specific regex pattern in all cells in which it occurs.
        var searchRegex = new Regex("Luna (\\d{2})");
        worksheet.Cells.ReplaceText(searchRegex, "REPLACED $1");

        workbook.Save("FoundAndReplaced.xlsx");
    }
}
