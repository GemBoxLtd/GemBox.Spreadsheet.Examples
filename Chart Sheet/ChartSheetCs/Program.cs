using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();

        int numberOfEmployees = 4;

        ExcelWorksheet ws1 = ef.Worksheets.Add("SourceSheet");

        // Add data which is used by the Excel chart.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; i++)
        {
            ws1.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());
            ws1.Cells[i + 1, 1].SetValue(random.Next(1000, 5000));
        }

        // Set header row and formatting.
        ws1.Cells[0, 0].Value = "Name";
        ws1.Cells[0, 1].Value = "Salary";
        ws1.Cells[0, 0].Style.Font.Weight = ws1.Cells[0, 1].Style.Font.Weight = ExcelFont.BoldWeight;
        ws1.Columns[0].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
        ws1.Columns[1].Style.NumberFormat = "\"$\"#,##0";

        // Create Excel chart sheet.
        ExcelWorksheet ws2 = ef.Worksheets.Add(SheetType.Chart, "ChartSheet");

        // Create Excel chart and select data for it.
        // You cannot set the size of the chart area when the chart is located on a chart sheet, it will snap to maximum size on the chart sheet.
        var chart = ws2.Charts.Add(ChartType.Bar, 0, 0, 0, 0, LengthUnit.Centimeter);
        chart.SelectData(ws1.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, 1), true);

        ef.Save("Chart Sheet.xlsx");
    }
}
