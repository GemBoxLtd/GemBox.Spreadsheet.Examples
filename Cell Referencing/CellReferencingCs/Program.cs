using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Cell Referencing");

        ws.Cells[0].Value = "Cell referencing examples:";

        ws.Cells["B2"].Value = "Cell B2.";
        ws.Cells[6, 0].Value = "Cell in row 7 and column A.";

        ws.Rows[2].Cells[0].Value = "Cell in row 3 and column A.";
        ws.Rows["4"].Cells["B"].Value = "Cell in row 4 and column B.";

        ws.Columns[2].Cells[4].Value = "Cell in column C and row 5.";
        ws.Columns["AA"].Cells["6"].Value = "Cell in AA column and row 6.";

        // Referencing Excel row's cell range.
        CellRange cr = ws.Rows[7].Cells;

        cr[0].Value = cr.IndexingMode.ToString();
        cr[3].Value = "D8";
        cr["B"].Value = "B8";

        // Referencing Excel column's cell range.
        cr = ws.Columns[7].Cells;

        cr[0].Value = cr.IndexingMode.ToString();
        cr[2].Value = "H3";
        cr["5"].Value = "H5";

        // Referencing arbitrary Excel cell range.
        cr = ws.Cells.GetSubrange("I2", "L8");
        cr.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromArgb(0, 0, 128), LineStyle.Dashed);

        cr["J7"].Value = cr.IndexingMode.ToString();
        cr[0, 0].Value = "I2";
        cr["J3"].Value = "J3";
        cr[4].Value = "I3"; // Cell range width is 4 (I J K L).

        // Autofit columns and some print options (for better look when exporting to pdf, xps and printing).
        var columnCount = ws.CalculateMaxUsedColumns();
        for (int i = 0; i < columnCount; i++)
            ws.Columns[i].AutoFit();

        ws.PrintOptions.PrintGridlines = true;
        ws.PrintOptions.PrintHeadings = true;

        ef.Save("Cell Referencing.xlsx");
    }
}
