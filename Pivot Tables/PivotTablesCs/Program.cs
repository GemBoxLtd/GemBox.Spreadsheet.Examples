using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.PivotTables;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();

        ExcelWorksheet worksheet1 = workbook.Worksheets.Add("SourceSheet");

        // Specify sheet formatting.
        worksheet1.Rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
        worksheet1.Columns[0].SetWidth(3, LengthUnit.Centimeter);
        worksheet1.Columns[1].SetWidth(3, LengthUnit.Centimeter);
        worksheet1.Columns[2].SetWidth(3, LengthUnit.Centimeter);
        worksheet1.Columns[3].SetWidth(3, LengthUnit.Centimeter);
        worksheet1.Columns[3].Style.NumberFormat = "[$$-409]#,##0.00";

        var cells = worksheet1.Cells;

        // Specify header row.
        cells[0, 0].Value = "Departments";
        cells[0, 1].Value = "Names";
        cells[0, 2].Value = "Years of Service";
        cells[0, 3].Value = "Salaries";

        // Insert random data to sheet.
        var random = new Random();
        var departments = new string[] { "Legal", "Marketing", "Finance", "Planning", "Purchasing" };
        var names = new string[] { "John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat" };
        var years = new string[] { "1-10", "11-20", "21-30", "over 30" };
        for (int i = 0; i < 100; ++i)
        {
            cells[i + 1, 0].Value = departments[random.Next(departments.Length)];
            cells[i + 1, 1].Value = names[random.Next(names.Length)] + ' ' + (i + 1).ToString();
            cells[i + 1, 2].Value = years[random.Next(years.Length)];
            cells[i + 1, 3].SetValue(random.Next(10, 101) * 100);
        }

        // Create pivot cache from cell range "SourceSheet!A1:D100".
        var cache = workbook.PivotCaches.AddWorksheetSource("SourceSheet!A1:D100");

        // Create new sheet for pivot table.
        var worksheet2 = workbook.Worksheets.Add("PivotSheet");

        // Create pivot table "Company Profile" using the specified pivot cache and add it to the worksheet at the cell location 'A1'.
        var table = worksheet2.PivotTables.Add(cache, "Company Profile", "A1");

        // Aggregate 'Names' values into count value and show it as a percentage of row.
        var field = table.DataFields.Add("Names");
        field.Function = PivotFieldCalculationType.Count;
        field.ShowDataAs = PivotFieldDisplayFormat.PercentageOfRow;
        field.Name = "% of Empl.";

        // Aggregate 'Salaries' values into average value.
        field = table.DataFields.Add("Salaries");
        field.Function = PivotFieldCalculationType.Average;
        field.Name = "Avg. Salary";
        field.NumberFormat = "[$$-409]#,##0.00";

        // Group rows into 'Departments'.
        table.RowFields.Add("Departments");

        // Group columns first into 'Years of Service' and then into 'Values' (count 'Names' and average 'Salaries').
        table.ColumnFields.Add("Years of Service");
        table.ColumnFields.Add(table.DataPivotField);

        // Specify the string to be displayed in row and column header.
        table.RowHeaderCaption = "Departments";
        table.ColumnHeaderCaption = "Years of Service";

        // Do not show grand totals for rows.
        table.RowGrandTotals = false;

        // Set pivot table style.
        table.BuiltInStyle = BuiltInPivotStyleName.PivotStyleMedium7;

        workbook.Save("Pivot Tables.xlsx");
    }
}