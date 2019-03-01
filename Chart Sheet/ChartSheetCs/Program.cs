using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();

        int numberOfEmployees = 4;

        var worksheet = workbook.Worksheets.Add("SourceSheet");

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

        // Create Excel chart sheet.
        var chartsheet = workbook.Worksheets.Add(SheetType.Chart, "ChartSheet");

        // Create Excel chart and select data for it.
        // You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        var chart = chartsheet.Charts.Add(ChartType.Bar, 0, 0, 0, 0, LengthUnit.Centimeter);
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        workbook.Save("Chart Sheet.xlsx");
    }
}