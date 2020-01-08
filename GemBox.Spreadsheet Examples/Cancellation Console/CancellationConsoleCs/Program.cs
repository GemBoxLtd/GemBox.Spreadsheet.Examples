using System;
using System.Diagnostics;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        // Use Trial Mode
        SpreadsheetInfo.FreeLimitReached += (eventSender, args) => args.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        // Create workbook
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("sheet");
        for (int i = 0; i < 1000000; i++)
            worksheet.Cells[i, 0].Value = i;

        var stopwatch = new Stopwatch();
        stopwatch.Start();

        // Create save options
        var saveOptions = new XlsxSaveOptions();
        saveOptions.ProgressChanged += (sender, args) =>
        {
            // Cancel operation after five seconds
            if (stopwatch.Elapsed.Seconds >= 5)
                args.CancelOperation();
        };

        try
        {
            workbook.Save("Cancellation.xlsx", saveOptions);
            Console.WriteLine("Operation fully finished");
        } 
        catch(OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled");
        }
    }
}