using System;
using System.Data;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var dataTable = new DataTable();

        // Depending on the format of the input file, you need to change this:
        dataTable.Columns.Add("FirstColumn", typeof(string));
        dataTable.Columns.Add("SecondColumn", typeof(string));

        // Select the first worksheet from the file.
        var worksheet = workbook.Worksheets[0];

        var options = new ExtractToDataTableOptions(0, 0, 10);
        options.ExtractDataOptions = ExtractDataOptions.StopAtFirstEmptyRow;
        options.ExcelCellToDataTableCellConverting += (sender, e) =>
        {
            if (!e.IsDataTableValueValid)
            {
                // GemBox.Spreadsheet doesn't automatically convert numbers to strings in ExtractToDataTable() method because of culture issues; 
                // someone would expect the number 12.4 as "12.4" and someone else as "12,4".
                e.DataTableValue = e.ExcelCell.Value == null ? null : e.ExcelCell.Value.ToString();
                e.Action = ExtractDataEventAction.Continue;
            }
        };

        // Extract the data from an Excel worksheet to the DataTable.
        // Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
        worksheet.ExtractToDataTable(dataTable, options);

        // Write DataTable content.
        var sb = new StringBuilder();
        sb.AppendLine("DataTable content:");
        foreach (DataRow row in dataTable.Rows)
        {
            sb.AppendFormat("{0}    {1}", row[0], row[1]);
            sb.AppendLine();
        }

        Console.WriteLine(sb.ToString());
    }
}