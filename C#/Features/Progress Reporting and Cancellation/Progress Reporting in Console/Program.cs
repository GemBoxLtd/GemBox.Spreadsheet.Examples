using GemBox.Spreadsheet;
using System;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Console.WriteLine("Creating file");

        // Create large workbook.
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("sheet");
        for (int i = 0; i < 1000000; i++)
            worksheet.Cells[i, 0].Value = i;

        // Create save options.
        var saveOptions = new XlsxSaveOptions();
        saveOptions.ProgressChanged += (eventSender, args) =>
        {
            Console.WriteLine($"Progress changed - {args.ProgressPercentage}%");
        };

        // Save file.
        workbook.Save("file.xlsx", saveOptions);
    }
}
