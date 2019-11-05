using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("NumberFormat.xlsx");

        var worksheet = workbook.Worksheets[0];

        worksheet.Cells[0, 2].Value = "ExcelCell.Value";
        worksheet.Columns[2].Style.NumberFormat = "@";

        worksheet.Cells[0, 3].Value = "CellStyle.NumberFormat";
        worksheet.Columns[3].Style.NumberFormat = "@";

        worksheet.Cells[0, 4].Value = "ExcelCell.GetFormattedValue()";
        worksheet.Columns[4].Style.NumberFormat = "@";

        for (int i = 1; i < worksheet.Rows.Count; i++)
        {
            var sourceCell = worksheet.Cells[i, 0];

            worksheet.Cells[i, 2].Value = sourceCell.Value?.ToString();
            worksheet.Cells[i, 3].Value = sourceCell.Style.NumberFormat;
            worksheet.Cells[i, 4].Value = sourceCell.GetFormattedValue();
        }

        // Set column widths.
        var columnWidths = new double[] { 192, double.NaN, 122, 236, 200 };
        for (int i = 0; i < columnWidths.Length; i++)
            if (!double.IsNaN(columnWidths[i]))
                worksheet.Columns[i].SetWidth(columnWidths[i], LengthUnit.Pixel);

        workbook.Save("Number Format.xlsx");
    }
}
