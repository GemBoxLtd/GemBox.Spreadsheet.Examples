using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // Load Excel workbook from file's path.
        ExcelFile workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        // Iterate through all worksheets in a workbook.
        foreach (ExcelWorksheet worksheet in workbook.Worksheets)
        {
            // Display sheet's name.
            Console.WriteLine("{1} {0} {1}\n", worksheet.Name, new string('#', 30));

            // Iterate through all rows in a worksheet.
            foreach (ExcelRow row in worksheet.Rows)
            {
                // Iterate through all allocated cells in a row.
                foreach (ExcelCell cell in row.AllocatedCells)
                {
                    // Read cell's data.
                    string value = cell.Value?.ToString() ?? "EMPTY";

                    // Display cell's value and type.
                    value = value.Length > 15 ? value.Remove(15) + "…" : value;
                    Console.Write($"{value} [{cell.ValueType}]".PadRight(30));
                }

                Console.WriteLine();
            }
        }
    }

    static void Example2()
    {
        ExcelFile workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
        {
            // Get Excel worksheet using zero-based index.
            ExcelWorksheet worksheet = workbook.Worksheets[sheetIndex];
            Console.WriteLine($"Sheet name: \"{worksheet.Name}\"");
            Console.WriteLine($"Sheet index: {worksheet.Index}\n");

            for (int rowIndex = 0; rowIndex < worksheet.Rows.Count; rowIndex++)
            {
                // Get Excel row using zero-based index.
                ExcelRow row = worksheet.Rows[rowIndex];
                Console.WriteLine($"Row name: \"{row.Name}\"");
                Console.WriteLine($"Row index: {row.Index}");

                Console.Write("Cell names:");
                for (int columnIndex = 0; columnIndex < row.AllocatedCells.Count; columnIndex++)
                {
                    // Get Excel cell using zero-based index.
                    ExcelCell cell = row.Cells[columnIndex];
                    Console.Write($" \"{cell.Name}\",");
                }
                Console.WriteLine("\n");
            }
        }
    }

    static void Example3()
    {
        ExcelFile workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        foreach (ExcelWorksheet worksheet in workbook.Worksheets)
        {
            CellRangeEnumerator enumerator = worksheet.Cells.GetReadEnumerator();
            while (enumerator.MoveNext())
            {
                ExcelCell cell = enumerator.Current;
                Console.WriteLine($"Cell \"{cell.Name}\" [{cell.Row.Index}, {cell.Column.Index}]: {cell.Value}");
            }
        }
    }
}