using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Filtering");

        int rowCount = 149;

        // Specify sheet formatting.
        worksheet.Rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet.Columns[0].SetWidth(3, LengthUnit.Centimeter);
        worksheet.Columns[1].SetWidth(3, LengthUnit.Centimeter);
        worksheet.Columns[2].SetWidth(3, LengthUnit.Centimeter);
        worksheet.Columns[2].Style.NumberFormat = "[$$-409]#,##0.00";
        worksheet.Columns[3].SetWidth(3, LengthUnit.Centimeter);
        worksheet.Columns[3].Style.NumberFormat = "yyyy-mm-dd";

        var cells = worksheet.Cells;

        // Specify header row.
        cells[0, 0].Value = "Departments";
        cells[0, 1].Value = "Names";
        cells[0, 2].Value = "Salaries";
        cells[0, 3].Value = "Deadlines";

        // Insert random data to sheet.
        var random = new Random();
        var departments = new string[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        for (int i = 0; i < rowCount; ++i)
        {
            cells[i + 1, 0].Value = departments[random.Next(departments.Length)];
            cells[i + 1, 1].Value = names[random.Next(names.Length)] + ' ' + (i + 1).ToString();
            cells[i + 1, 2].SetValue(random.Next(10, 101) * 100);
            cells[i + 1, 3].SetValue(DateTime.Now.AddDays(random.Next(-1, 2)));
        }

        // Specify range which will be filtered.
        var filterRange = worksheet.Cells.GetSubrangeAbsolute(0, 0, rowCount, 3);

        // Show only rows which satisfy following conditions:
        // - 'Departments' value is either "Legal" or "Marketing" or "Finance" and
        // - 'Names' value contains letter 'e' and
        // - 'Salaries' value is in the top 20 percent of all 'Salaries' values and
        // - 'Deadlines' value is today's date.
        // Shown rows are then sorted by 'Salaries' values in the descending order.
        filterRange.Filter().
            ByValues(0, "Legal", "Marketing", "Finance").
            ByCustom(1, FilterOperator.Equal, "*e*").
            ByTop10(2, true, true, 20).
            ByDynamic(3, DynamicFilterType.Today).
            SortBy(2, true).
            Apply();

        workbook.Save("Filtering.xlsx");
    }
}