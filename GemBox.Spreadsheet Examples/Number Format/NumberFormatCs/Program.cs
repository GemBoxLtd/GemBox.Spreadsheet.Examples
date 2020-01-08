using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    public static void Example1()
    {
        var workbook = ExcelFile.Load("NumberFormat.xlsx");

        var worksheet = workbook.Worksheets[0];

        worksheet.Cells[0, 2].Value = "ExcelCell.Value";
        worksheet.Columns[2].Style.NumberFormat = "@";

        worksheet.Cells[0, 3].Value = "CellStyle.NumberFormat";
        worksheet.Columns[3].Style.NumberFormat = "@";

        worksheet.Cells[0, 4].Value = "ExcelCell.GetFormattedValue()";
        worksheet.Columns[4].Style.NumberFormat = "@";

        for (int i = 1; i < worksheet.Rows.Count; i++)
        {
            var sourceCell = worksheet.Cells[i, 0];

            worksheet.Cells[i, 2].Value = sourceCell.Value?.ToString();
            worksheet.Cells[i, 3].Value = sourceCell.Style.NumberFormat;
            worksheet.Cells[i, 4].Value = sourceCell.GetFormattedValue();
        }

        // Set column widths.
        var columnWidths = new double[] { 192, double.NaN, 122, 236, 200 };
        for (int i = 0; i < columnWidths.Length; i++)
            if (!double.IsNaN(columnWidths[i]))
                worksheet.Columns[i].SetWidth(columnWidths[i], LengthUnit.Pixel);

        workbook.Save("Number Format.xlsx");
    }

    public static void Example2()
    {
        var workbook = new ExcelFile();

        var worksheet = workbook.Worksheets.Add("sheet");

        worksheet.Columns[0].SetWidth(200, LengthUnit.Pixel);

        // Show the value as a number with two decimal places and thousands separator.
        worksheet.Cells[0, 0].Style.NumberFormat
            = NumberFormatBuilder.Number(2, useThousandsSeparator: true);
        worksheet.Cells[0, 0].Value = 2500.333;

        // Show the value in Euros and display negative values in parentheses.
        worksheet.Cells[1, 0].Style.NumberFormat
            = NumberFormatBuilder.Currency("€", 2, useParenthesesToDisplayNegativeValue: true);
        worksheet.Cells[1, 0].Value = -50;

        // Show the value in accounting format with three decimal places.
        worksheet.Cells[2, 0].Style.NumberFormat
            = NumberFormatBuilder.Accounting(3, currencySymbol: "$");
        worksheet.Cells[2, 0].Value = -50;

        // Show the value in ISO 8061 date format.
        worksheet.Cells[3, 0].Style.NumberFormat = NumberFormatBuilder.DateTimeIso8061();
        worksheet.Cells[3, 0].Value = DateTime.Now;

        // Show the value as percentage.
        worksheet.Cells[4, 0].Style.NumberFormat = NumberFormatBuilder.Percentage(2);
        worksheet.Cells[4, 0].Value = 1 / 3d;

        // Show the value as fraction with 100 as a denominator.
        worksheet.Cells[5, 0].Style.NumberFormat = NumberFormatBuilder.FractionWithPreciseDenominator(100);
        worksheet.Cells[5, 0].Value = 1 / 3d;

        // Show the value in scientific notation using two decimal places.
        worksheet.Cells[6, 0].Style.NumberFormat = NumberFormatBuilder.Scientific(2);
        worksheet.Cells[6, 0].Value = Math.Pow(Math.PI, 10);

        workbook.Save("Number Format Builder.docx");
    }
}
