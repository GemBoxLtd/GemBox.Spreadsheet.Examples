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

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        // Select the first worksheet from the file.
        ExcelWorksheet ws = ef.Worksheets[0];

        // Create DataTable from an Excel worksheet.
        DataTable dataTable = ws.CreateDataTable(new CreateDataTableOptions()
        {
            ColumnHeaders = true,
            StartRow = 1,
            NumberOfColumns = 5,
            NumberOfRows = ws.Rows.Count - 1,
            Resolution = ColumnTypeResolution.AutoPreferStringCurrentCulture
        });

        // Write DataTable content
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("DataTable content:");
        foreach (DataRow row in dataTable.Rows)
        {
            sb.AppendFormat("{0}\t{1}\t{2}\t{3}\t{4}", row[0], row[1], row[2], row[3], row[4]);
            sb.AppendLine();
        }

        Console.WriteLine(sb.ToString());
    }
}
