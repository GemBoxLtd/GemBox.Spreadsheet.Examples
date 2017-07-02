using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("ChartTemplate.xlsx");

        int numberOfEmployees = 4;

        var ws = ef.Worksheets[0];

        // Update named ranges 'Names' and 'Salaries' which are used by preserved chart.
        ws.NamedRanges["Names"].Range = ws.Cells.GetSubrangeAbsolute(1, 0, numberOfEmployees, 0);
        ws.NamedRanges["Salaries"].Range = ws.Cells.GetSubrangeAbsolute(1, 1, numberOfEmployees, 1);

        // Add data which is used by preserved chart through named ranges 'Names' and 'Salaries'.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            ws.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());
            ws.Cells[i + 1, 1].SetValue(random.Next(1000, 5000));
        }

        ef.Save("Preservation.xlsx");
    }
}
