using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfEmployees = 4;
        int numberOfYears = 4;

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Chart");

        // Create chart and select data for it.
        var chart = (ColumnChart)worksheet.Charts.Add(ChartType.Column, "B7", "O27");
        chart.SelectData(worksheet.Cells.GetSubrangeAbsolute(0, 0, numberOfEmployees, numberOfYears));

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
            worksheet.Cells[i + 1, 0].Value = names[i % names.Length] + (i < names.Length ? string.Empty : ' ' + (i / names.Length + 1).ToString());

            for (int j = 0; j < numberOfYears; ++j)
                worksheet.Cells[i + 1, j + 1].SetValue(random.Next(1000, 5000));
        }

        // Set header row and formatting.
        worksheet.Cells[0, 0].Value = "Name";
        worksheet.Cells[0, 0].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet.Columns[0].Width = (int)LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.ZeroCharacterWidth256thPart);
        for (int i = 0, startYear = DateTime.Now.Year - numberOfYears; i < numberOfYears; ++i)
        {
            worksheet.Cells[0, i + 1].SetValue(startYear + i);
            worksheet.Cells[0, i + 1].Style.Font.Weight = ExcelFont.BoldWeight;
            worksheet.Cells[0, i + 1].Style.NumberFormat = "General";
            worksheet.Columns[i + 1].Style.NumberFormat = "\"$\"#,##0";
        }

        // Make entire sheet print horizontally centered on a single page with headings and gridlines.
        var printOptions = worksheet.PrintOptions;
        printOptions.HorizontalCentered = printOptions.PrintHeadings = printOptions.PrintGridlines = true;
        printOptions.FitWorksheetWidthToPages = printOptions.FitWorksheetHeightToPages = 1;

        workbook.Save("Chart Components.xlsx");
    }
}