using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;
using System;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Chart");

        int numberOfYears = 4;

        // Add data which is used by the chart.
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "John Doe";
        worksheet.Cells["A3"].Value = "Fred Nurk";
        worksheet.Cells["A4"].Value = "Hans Meier";
        worksheet.Cells["A5"].Value = "Ivan Horvat";

        // Generate column titles.
        for (int i = 0; i < numberOfYears; i++)
            worksheet.Cells[0, i + 1].Value = DateTime.Now.Year - numberOfYears + i;

        var random = new Random();
        var range = worksheet.Cells.GetSubrangeAbsolute(1, 1, 4, numberOfYears);

        // Fill the values.
        foreach (var cell in range)
        {
            cell.SetValue(random.Next(1000, 5000));
            cell.Style.NumberFormat = "\"$\"#,##0";
        }

        // Set header row and formatting.
        worksheet.Rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet.Rows[0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        worksheet.Columns[0].SetWidth(3, LengthUnit.Centimeter);

        // Create chart and select data for it.
        var chart = worksheet.Charts.Add<ColumnChart>(ChartGrouping.Clustered, "B7", "O27");
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, 4, numberOfYears));

        // Set chart title.
        chart.Title.Text = "Column Chart";

        // Set chart legend.
        chart.Legend.IsVisible = true;
        chart.Legend.Position = ChartLegendPosition.Right;

        // Set axis titles.
        chart.Axes.Horizontal.Title.Text = "Years";
        chart.Axes.Vertical.Title.Text = "Salaries";

        // Set value axis scaling, units, gridlines and tick marks.
        var valueAxis = chart.Axes.VerticalValue;
        valueAxis.Minimum = 0;
        valueAxis.Maximum = 6000;
        valueAxis.MajorUnit = 1000;
        valueAxis.MinorUnit = 500;
        valueAxis.MajorGridlines.IsVisible = true;
        valueAxis.MinorGridlines.IsVisible = true;
        valueAxis.MajorTickMarkType = TickMarkType.Outside;
        valueAxis.MinorTickMarkType = TickMarkType.Cross;

        // Make entire sheet print horizontally centered on a single page with headings and gridlines.
        var printOptions = worksheet.PrintOptions;
        printOptions.HorizontalCentered = true;
        printOptions.PrintHeadings = true;
        printOptions.PrintGridlines = true;
        printOptions.FitWorksheetWidthToPages = 1;
        printOptions.FitWorksheetHeightToPages = 1;

        workbook.Save("Chart Components.xlsx");
    }
}
