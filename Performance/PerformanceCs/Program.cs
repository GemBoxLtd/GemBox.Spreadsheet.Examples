using System;
using System.Diagnostics;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If sample exceeds Free version limitations then continue as Trial version:
        // https://www.gemboxsoftware.com/spreadsheet/help/html/Evaluation_and_Licensing.htm
        SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        int rowCount = 100000;
        int columnCount = 10;
        string fileFormat = "XLSX";

        Console.WriteLine("Performance sample:");
        Console.WriteLine();
        Console.WriteLine("Row count: " + rowCount);
        Console.WriteLine("Column count: " + columnCount);
        Console.WriteLine("File format: " + fileFormat);
        Console.WriteLine();

        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Performance");

        for (int row = 0; row < rowCount; row++)
            for (int column = 0; column < columnCount; column++)
                ws.Cells[row, column].Value = row.ToString() + "_" + column;

        Console.WriteLine("Generate file (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        int cellsCount = 0;
        foreach (var row in ws.Rows)
            foreach (var cell in row.AllocatedCells)
                ++cellsCount;

        Console.WriteLine("Iterate through " + cellsCount + " cells (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        ef.Save("Report." + fileFormat.ToLower());

        Console.WriteLine("Save file (seconds): " + stopwatch.Elapsed.TotalSeconds);
    }
}