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
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formats");

        worksheet.Rows[0].Style = workbook.Styles[BuiltInCellStyleName.Heading1];
        worksheet.Columns[0].Width = 25 * 256;
        worksheet.Columns[1].Width = 35 * 256;
        worksheet.Columns[2].Width = 25 * 256;

        worksheet.Cells[0, 0].Value = "Value & Format";
        worksheet.Cells[0, 1].Value = "Format";
        worksheet.Cells[0, 2].Value = "Type";

        // Sample data with values and formats.
        var data = new (object Value, string Format)[]
        {
            (1.23, "0"),
            (1.23, "0.00"),
            (1.2345, "0.000"),
            (-2.345, @"0.00_);[Red]\(0.00\)"),
            (2.34, @"\$#,##0.00"),
            (2345.67, @"#,##0.00\ [$€-1]"),
            (new DateTime(2012, 11, 9, 0, 0, 0), @"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"),
            (new DateTime(2012, 12, 5, 0, 0, 0), @"[$-409]mmmm\ d\,\ yyyy;@"),
            (new DateTime(2012, 8, 10, 0, 0, 0), @"yyyy\-mm\-dd\ \(dddd\)"),
            (new DateTime(2012, 8, 12, 0, 13, 0), @"[$-409]m/d/yy\ h:mm\ AM/PM;@"),
            (new DateTime(2012, 8, 1, 21, 10, 0), @"[$-409]h:mm\ AM/PM;@"),
            (new DateTime(1900, 1, 1, 6, 45, 30), "[h]:mm:ss"),
            (0.0123, "0%"),
            (0.0123, "0.00%"),
            (120000, "0.00E+00"),
            (1.25, @"# ?/?"),
            (1.25, @"#\ ?/100"),
            ("Sample text", "@")
        };

        for (int i = 0; i < data.Length; i++)
        {
            var item = data[i];

            // Write value and set number format to a cell.
            worksheet.Cells[i + 1, 0].Value = item.Value;
            worksheet.Cells[i + 1, 0].Style.NumberFormat = item.Format;

            // Write number format as cell's value.
            worksheet.Cells[i + 1, 1].Value = item.Format;

            // Write data type as cell's value.
            worksheet.Cells[i + 1, 2].Value = item.Value.GetType().ToString();
        }

        workbook.Save("Number Formats.xlsx");
    }

    public static void Example2()
    {
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("sheet");
        worksheet.Columns[0].SetWidth(200, LengthUnit.Pixel);

        // Show the value as a number with two decimal places and thousands separator.
        worksheet.Cells[0, 0].Style.NumberFormat =
            NumberFormatBuilder.Number(2, useThousandsSeparator: true);
        worksheet.Cells[0, 0].Value = 2500.333;

        // Show the value in Euros and display negative values in parentheses.
        worksheet.Cells[1, 0].Style.NumberFormat =
            NumberFormatBuilder.Currency("€", 2, useParenthesesToDisplayNegativeValue: true);
        worksheet.Cells[1, 0].Value = -50;

        // Show the value in accounting format with three decimal places.
        worksheet.Cells[2, 0].Style.NumberFormat =
            NumberFormatBuilder.Accounting(3, currencySymbol: "$");
        worksheet.Cells[2, 0].Value = -50;

        // Show the value in ISO 8061 date format.
        worksheet.Cells[3, 0].Style.NumberFormat =
            NumberFormatBuilder.DateTimeIso8061();
        worksheet.Cells[3, 0].Value = DateTime.Now;

        // Show the value as percentage.
        worksheet.Cells[4, 0].Style.NumberFormat =
            NumberFormatBuilder.Percentage(2);
        worksheet.Cells[4, 0].Value = 1 / 3d;

        // Show the value as fraction with 100 as a denominator.
        worksheet.Cells[5, 0].Style.NumberFormat =
            NumberFormatBuilder.FractionWithPreciseDenominator(100);
        worksheet.Cells[5, 0].Value = 1 / 3d;

        // Show the value in scientific notation using two decimal places.
        worksheet.Cells[6, 0].Style.NumberFormat =
            NumberFormatBuilder.Scientific(2);
        worksheet.Cells[6, 0].Value = Math.Pow(Math.PI, 10);

        workbook.Save("Number Format Builder.xlsx");
    }
}
