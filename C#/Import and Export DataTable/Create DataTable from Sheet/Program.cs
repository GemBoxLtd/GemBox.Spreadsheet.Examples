using System;
using System.Data;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
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