using System;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        StringBuilder sb = new StringBuilder();

        // Iterate through all worksheets in an Excel workbook.
        foreach (ExcelWorksheet sheet in ef.Worksheets)
        {
            sb.AppendLine();
            sb.AppendFormat("{0} {1} {0}", new string('-', 25), sheet.Name);

            // Iterate through all rows in an Excel worksheet.
            foreach (ExcelRow row in sheet.Rows)
            {
                sb.AppendLine();

                // Iterate through all allocated cells in an Excel row.
                foreach (ExcelCell cell in row.AllocatedCells)
                {
                    if (cell.ValueType != CellValueType.Null)
                        sb.Append(string.Format("{0} [{1}]", cell.Value, cell.ValueType).PadRight(25));
                    else
                        sb.Append(new string(' ', 25));
                }
            }
        }

        Console.WriteLine(sb.ToString());
    }
}
