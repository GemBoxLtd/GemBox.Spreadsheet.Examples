using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Chart");

        int numberOfEmployees = 4;

        // Create Excel chart and select data for it.
        var chart = ws.Charts.Add(ChartType.Bar, "D2", "M25");
        chart.SelectData(ws.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        // Add data which is used by the Excel chart.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            ws.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());
            ws.Cells[i + 1, 1].SetValue(random.Next(1000, 5000));
        }

        // Set header row and formatting.
        ws.Cells[0, 0].Value = "Name";
        ws.Cells[0, 1].Value = "Salary";
        ws.Cells[0, 0].Style.Font.Weight = ws.Cells[0, 1].Style.Font.Weight = ExcelFont.BoldWeight;
        ws.Columns[0].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
        ws.Columns[1].Style.NumberFormat = "\"$\"#,##0";

        // Make entire sheet print on a single page.
        ws.PrintOptions.FitWorksheetWidthToPages = ws.PrintOptions.FitWorksheetHeightToPages = 1;

        ef.Save("Chart.xlsx");
    }
}
