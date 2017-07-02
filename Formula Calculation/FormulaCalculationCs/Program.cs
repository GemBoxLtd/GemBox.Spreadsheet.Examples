using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Formula Calculation");

        // Some formatting.
        ExcelRow row = ws.Rows[0];
        row.Style.Font.Weight = ExcelFont.BoldWeight;

        ExcelColumn col = ws.Columns[0];
        col.SetWidth(250, LengthUnit.Pixel);
        col.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
        col = ws.Columns[1];
        col.SetWidth(250, LengthUnit.Pixel);
        col.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

        // Use first row for column headers.
        ws.Cells["A1"].Value = "Formula";
        ws.Cells["B1"].Value = "Calculated value";

        // Enter some Excel formulas as text in first column.
        ws.Cells["A2"].Value = "=1 + 1";
        ws.Cells["A3"].Value = "=3 * (2 - 8)";
        ws.Cells["A4"].Value = "=3 + ABS(B3)";
        ws.Cells["A5"].Value = "=B4 > 15";
        ws.Cells["A6"].Value = "=IF(B5, \"Hello world\", \"World hello\")";
        ws.Cells["A7"].Value = "=B6 & \" example\"";
        ws.Cells["A8"].Value = "=CODE(RIGHT(B7))";
        ws.Cells["A9"].Value = "=POWER(B8, 3) * 0.45%";
        ws.Cells["A10"].Value = "=SIGN(B9)";
        ws.Cells["A11"].Value = "=SUM(B2:B10)";

        // Set text from first column as second row cell's formula.
        int rowIndex = 1;
        while (ws.Cells[rowIndex, 0].ValueType != CellValueType.Null)
            ws.Cells[rowIndex, 1].Formula = ws.Cells[rowIndex++, 0].StringValue;

        // GemBox.Spreadsheet supports single Excel cell calculation, ...
        ws.Cells["B2"].Calculate();

        // ... Excel worksheet calculation,
        ws.Calculate();

        // ... and whole Excel file calculation.
        ws.Parent.Calculate();

        ef.Save("Formula Calculation.xlsx");
    }
}
