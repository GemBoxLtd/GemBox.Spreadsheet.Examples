using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfEmployees = 4;
        int numberOfYears = 4;

        var ef = new ExcelFile();
        var ws = ef.Worksheets.Add("Chart");

        // Create chart and select data for it.
        var chart = (ColumnChart)ws.Charts.Add(ChartType.Column, "B7", "O27");
        chart.SelectData(ws.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, numberOfYears));

        // Set chart title.
        chart.Title.Text = "Clustered Column Chart";

        // Set axis titles.
        chart.Axes.Horizontal.Title.Text = "Years";
        chart.Axes.Vertical.Title.Text = "Salaries";

        // For all charts (except Pie and Bar) value axis is vertical.
        var valueAxis = chart.Axes.VerticalValue;

        // Set value axis scaling, units, gridlines and tick marks.
        valueAxis.Minimum = 0;
        valueAxis.Maximum = 6000;
        valueAxis.MajorUnit = 1000;
        valueAxis.MinorUnit = 500;
        valueAxis.MajorGridlines.IsVisible = true;
        valueAxis.MinorGridlines.IsVisible = true;
        valueAxis.MajorTickMarkType = TickMarkType.Outside;
        valueAxis.MinorTickMarkType = TickMarkType.Cross;

        // Add data which is used by the chart.
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var random = new Random();
        for (int i = 0; i < numberOfEmployees; ++i)
        {
            ws.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());

            for (int j = 0; j < numberOfYears; ++j)
                ws.Cells[i + 1, j + 1].SetValue(random.Next(1000, 5000));
        }

        // Set header row and formatting.
        ws.Cells[0, 0].Value = "Name";
        ws.Cells[0, 0].Style.Font.Weight = ExcelFont.BoldWeight;
        ws.Columns[0].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
        for (int i = 0, startYear = DateTime.Now.Year - numberOfYears; i < numberOfYears; ++i)
        {
            ws.Cells[0, i + 1].SetValue(startYear + i);
            ws.Cells[0, i + 1].Style.Font.Weight = ExcelFont.BoldWeight;
            ws.Cells[0, i + 1].Style.NumberFormat = "General";
            ws.Columns[i + 1].Style.NumberFormat = "\"$\"#,##0";
        }

        // Make entire sheet print horizontally centered on a single page with headings and gridlines.
        var printOptions = ws.PrintOptions;
        printOptions.HorizontalCentered = printOptions.PrintHeadings = printOptions.PrintGridlines = true;
        printOptions.FitWorksheetWidthToPages = printOptions.FitWorksheetHeightToPages = 1;

        ef.Save("Chart Components.xlsx");
    }
}
