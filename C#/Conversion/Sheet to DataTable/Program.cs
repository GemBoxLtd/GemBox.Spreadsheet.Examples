using GemBox.Spreadsheet;
using System;
using System.Data;

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

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Create DataTable with specified columns.
        var dataTable = new DataTable();
        dataTable.Columns.Add("First_Column", typeof(string));
        dataTable.Columns.Add("Second_Column", typeof(string));
        dataTable.Columns.Add("Third_Column", typeof(int));
        dataTable.Columns.Add("Fourth_Column", typeof(double));

        // Select the first worksheet from the file.
        var worksheet = workbook.Worksheets[0];

        // Extract the data from an Excel worksheet to the DataTable.
        var options = new ExtractToDataTableOptions(0, 0, 20);
        options.ExcelCellToDataTableCellConverting += (sender, e) =>
        {
            if (!e.IsDataTableValueValid)
            {
                // Convert ExcelCell value to string.
                if (e.DataTableColumnType == typeof(string))
                    e.DataTableValue = e.ExcelCell.Value?.ToString();
                else
                    e.DataTableValue = DBNull.Value;
            }
        };
        worksheet.ExtractToDataTable(dataTable, options);

        // Write DataTable columns.
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(column.ColumnName.PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write($"[{column.DataType}]".PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(new string('-', column.ColumnName.Length).PadRight(20));
        Console.WriteLine();

        // Write DataTable rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object item in row.ItemArray)
            {
                string value = item.ToString();
                value = value.Length > 20 ? value.Remove(19) + "…" : value;
                Console.Write(value.PadRight(20));
            }
            Console.WriteLine();
        }
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Select the first worksheet from the file.
        var worksheet = workbook.Worksheets[0];

        // Create DataTable from an Excel worksheet.
        var dataTable = worksheet.CreateDataTable(new CreateDataTableOptions()
        {
            ColumnHeaders = true,
            StartRow = 1,
            NumberOfColumns = 5,
            NumberOfRows = worksheet.Rows.Count - 1,
            Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        });

        // Write DataTable columns.
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(column.ColumnName.PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write($"[{column.DataType}]".PadRight(20));
        Console.WriteLine();
        foreach (DataColumn column in dataTable.Columns)
            Console.Write(new string('-', column.ColumnName.Length).PadRight(20));
        Console.WriteLine();

        // Write DataTable rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object item in row.ItemArray)
            {
                string value = item.ToString();
                value = value.Length > 20 ? value.Remove(19) + "…" : value;
                Console.Write(value.PadRight(20));
            }
            Console.WriteLine();
        }
    }
}
