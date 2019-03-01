using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("ChartTemplate.xlsx");

        int numberOfEmployees = 4;

        var worksheet = workbook.Worksheets[0];

        // Update named ranges 'Names' and 'Salaries' which are used by preserved chart.
        worksheet.NamedRanges["Names"].Range = worksheet.Cells.GetSubrangeAbsolute(1, 0, numberOfEmployees, 0);
        worksheet.NamedRanges["Salaries"].Range = worksheet.Cells.GetSubrangeAbsolute(1, 1, numberOfEmployees, 1);

        // Add data which is used by preserved chart through named ranges 'Names' and 'Salaries'.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            worksheet.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());
            worksheet.Cells[i + 1, 1].SetValue(random.Next(1000, 5000));
        }

        workbook.Save("Preservation.xlsx");
    }
}