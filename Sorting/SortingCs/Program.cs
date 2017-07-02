using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Sorting");

        Random rnd = new Random();

        ws.Cells[0, 0].Value = "Sorted numbers";
        ws.Cells.GetSubrangeAbsolute(0, 0, 0, 1).Merged = true;
        for (int i = 1; i < 10; i++)
            ws.Cells[i, 0].SetValue(rnd.Next(1, 100));

        ws.Cells.GetSubrangeAbsolute(1, 0, 10, 0).Sort(false).By(0).Apply();

        ws.Cells[0, 2].Value = "Sorted strings";
        ws.Cells.GetSubrangeAbsolute(0, 2, 0, 3).Merged = true;
        ws.Cells[1, 2].Value = "John";
        ws.Cells[2, 2].Value = "Jennifer";
        ws.Cells[3, 2].Value = "Toby";
        ws.Cells[4, 2].Value = "Chloe";

        ws.Cells.GetSubrangeAbsolute(1, 2, 4, 2).Sort(false).By(0).Apply();

        ws.Cells[0, 4].Value = "Sorted by column E and after that by column F";
        ws.Cells.GetSubrangeAbsolute(0, 4, 0, 8).Merged = true;
        for (int i = 1; i < 10; i++)
        {
            ws.Cells[i, 4].SetValue(rnd.Next(1, 4));
            ws.Cells[i, 5].SetValue(rnd.Next(0, 10));
        }

        // Sort by column E ascending and then by column F descending.
        // These sort settings will be saved to output XLSX file because they are active (parameter value is true).
        ws.Cells.GetSubrangeAbsolute(1, 4, 10, 5).Sort(true).By(0).By(1, true).Apply();

        ef.Save("Sorting.xlsx");
    }
}
