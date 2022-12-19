using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sorting");

        var random = new Random();

        worksheet.Cells[0, 0].Value = "Sorted numbers";
        for (int i = 1; i < 10; i++)
            worksheet.Cells[i, 0].SetValue(random.Next(1, 100));

        worksheet.Cells.GetSubrangeAbsolute(1, 0, 10, 0).Sort(false).By(0).Apply();

        worksheet.Cells[0, 2].Value = "Sorted strings";
        worksheet.Cells[1, 2].Value = "John";
        worksheet.Cells[2, 2].Value = "Jennifer";
        worksheet.Cells[3, 2].Value = "Toby";
        worksheet.Cells[4, 2].Value = "Chloe";

        worksheet.Cells.GetSubrangeAbsolute(1, 2, 4, 2).Sort(false).By(0).Apply();

        worksheet.Cells[0, 4].Value = "Sorted by column E and after that by column F";
        for (int i = 1; i < 10; i++)
        {
            worksheet.Cells[i, 4].SetValue(random.Next(1, 4));
            worksheet.Cells[i, 5].SetValue(random.Next(0, 10));
        }

        // Sort by column E ascending and then by column F descending.
        // These sort settings will be saved to output XLSX file because they are active (parameter value is true).
        worksheet.Cells.GetSubrangeAbsolute(1, 4, 10, 5).Sort(true).By(0).By(1, true).Apply();

        workbook.Save("Sorting.xlsx");
    }
}