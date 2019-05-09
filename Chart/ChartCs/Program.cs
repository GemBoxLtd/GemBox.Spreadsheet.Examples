using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Chart");

        int numberOfEmployees = 4;

        // Create Excel chart and select data for it.
        var chart = worksheet.Charts.Add(ChartType.Bar, "D2", "M25");
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        // Add data which is used by the Excel chart.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            worksheet.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());
            worksheet.Cells[i + 1, 1].SetValue(random.Next(1000, 5000));
        }

        // Set header row and formatting.
        worksheet.Cells[0, 0].Value = "Name";
        worksheet.Cells[0, 1].Value = "Salary";
        worksheet.Cells[0, 0].Style.Font.Weight = worksheet.Cells[0, 1].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet.Columns[0].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
        worksheet.Columns[1].Style.NumberFormat = "\"$\"#,##0";

        // Make entire sheet print on a single page.
        worksheet.PrintOptions.FitWorksheetWidthToPages = worksheet.PrintOptions.FitWorksheetHeightToPages = 1;

        workbook.Save("Chart.xlsx");
    }
}