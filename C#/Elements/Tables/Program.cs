using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Tables;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");


        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Tables");

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

        workbook.Save("Tables.xlsx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("Tables.xlsx");
        var worksheet = workbook.Worksheets["Tables"];
        var table = worksheet.Tables["Table1"];

        // Remove existing table row.
        table.Rows.RemoveAt(0);

        // Update existing table row.
        var tableRow = table.Rows[0];
        tableRow.DataRange[0].Value = "Jane Updated";
        tableRow.DataRange[1].Value = 30;
        tableRow.DataRange[2].Value = 40.0;

        // Sample data for writing into a table.
        var data = new[]
        {
            new object[]{ "Fred Nurk", 22, 35.0 },
            new object[]{ "Hans Meier", 16, 20.0 },
            new object[]{ "Ivan Horvat", 24, 34.0 }
        };

        foreach (object[] items in data)
        {
            // Add new table row by adding cell values directly.
            tableRow = table.Rows.Add(items);
            tableRow.DataRange[3].Formula = "=Table1[Hours] * Table1[Price]";
        }

        table.Columns["Total"].Range.Calculate();

        workbook.Save("Tables Updated.xlsx");
    }
}
