using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var ef = new ExcelFile();
        var worksheet = ef.Worksheets.Add("Tables");

        // Add some data.
        var data = new object[5, 3]
        {
            { "Worker", "Hours", "Price" },
            { "John Doe", 25, 35.0 },
            { "Jane Doe", 27, 35.0 },
            { "Jack White", 18, 32.0 },
            { "George Black", 31, 35.0 }
        };

        for (int i = 0; i < 5; i++)
            for (int j = 0; j < 3; j++)
                worksheet.Cells[i, j].Value = data[i, j];

        // Set column widths.
        worksheet.Columns[0].SetWidth(100, LengthUnit.Pixel);
        worksheet.Columns[1].SetWidth(70, LengthUnit.Pixel);
        worksheet.Columns[2].SetWidth(70, LengthUnit.Pixel);
        worksheet.Columns[3].SetWidth(70, LengthUnit.Pixel);
        worksheet.Columns[2].Style.NumberFormat = "\"$\"#,##0.00";
        worksheet.Columns[3].Style.NumberFormat = "\"$\"#,##0.00";

        // Create table and enable totals row.
        var table = worksheet.Tables.Add("Table1", "A1:C5", true);
        table.HasTotalsRow = true;

        // Add new column.
        var column = table.Columns.Add();
        column.Name = "Total";

        // Populate column.
        foreach (var cell in column.DataRange)
            cell.Formula = "=Table1[Hours] * Table1[Price]";

        // Set totals row function for newly added column and calculate it.
        column.TotalsRowFunction = TotalsRowFunction.Sum;
        column.Range.Calculate();

        // Set table style.
        table.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2;

        ef.Save("Tables.xlsx");
    }
}