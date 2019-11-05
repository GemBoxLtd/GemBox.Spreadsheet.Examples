using System;
using System.Diagnostics;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If example exceeds Free version limitations then continue as Trial version:
        // https://www.gemboxsoftware.com/spreadsheet/help/html/Evaluation_and_Licensing.htm
        SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        int rowCount = 100000;
        int columnCount = 10;
        var fileFormat = "XLSX";

        Console.WriteLine("Performance example:");
        Console.WriteLine();
        Console.WriteLine("Row count: " + rowCount);
        Console.WriteLine("Column count: " + columnCount);
        Console.WriteLine("File format: " + fileFormat);
        Console.WriteLine();

        var stopwatch = new Stopwatch();
        stopwatch.Start();

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Performance");

        for (int row = 0; row < rowCount; row++)
            for (int column = 0; column < columnCount; column++)
                worksheet.Cells[row, column].Value = row.ToString() + "_" + column;

        Console.WriteLine("Generate file (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        int cellsCount = 0;
        foreach (var row in worksheet.Rows)
            foreach (var cell in row.AllocatedCells)
                ++cellsCount;

        Console.WriteLine("Iterate through " + cellsCount + " cells (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        workbook.Save("Report." + fileFormat.ToLower());

        Console.WriteLine("Save file (seconds): " + stopwatch.Elapsed.TotalSeconds);
    }
}