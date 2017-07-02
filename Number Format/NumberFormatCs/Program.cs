using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("NumberFormat.xlsx");

        var ws = ef.Worksheets[0];

        ws.Cells[0, 2].Value = "ExcelCell.Value";
        ws.Columns[2].Style.NumberFormat = "@";

        ws.Cells[0, 3].Value = "CellStyle.NumberFormat";
        ws.Columns[3].Style.NumberFormat = "@";

        ws.Cells[0, 4].Value = "ExcelCell.GetFormattedValue()";
        ws.Columns[4].Style.NumberFormat = "@";

        for (int i = 1; i < ws.Rows.Count; i++)
        {
            ExcelCell sourceCell = ws.Cells[i, 0];

            ws.Cells[i, 2].Value = sourceCell.Value == null ? null : sourceCell.Value.ToString();
            ws.Cells[i, 3].Value = sourceCell.Style.NumberFormat;
            ws.Cells[i, 4].Value = sourceCell.GetFormattedValue();
        }

        // Auto-fit columns
        for (int i = 0; i < 5; i++)
            ws.Columns[i].AutoFit();

        ef.Save("Number Format.xlsx");
    }
}
