using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("Template.xlsx");

        int workingDays = 8;

        DateTime startDate = DateTime.Now.AddDays(-workingDays);
        DateTime endDate = DateTime.Now;

        ExcelWorksheet ws = ef.Worksheets[0];

        // Find cells with placeholder text and set their values.
        int row, column;
        if (ws.Cells.FindText("[Company Name]", true, true, out row, out column))
            ws.Cells[row, column].Value = "ACME Corp";
        if (ws.Cells.FindText("[Company Address]", true, true, out row, out column))
            ws.Cells[row, column].Value = "240 Old Country Road, Springfield, IL";
        if (ws.Cells.FindText("[Start Date]", true, true, out row, out column))
            ws.Cells[row, column].Value = startDate;
        if (ws.Cells.FindText("[End Date]", true, true, out row, out column))
            ws.Cells[row, column].Value = endDate;

        // Copy template row.
        row = 17;
        ws.Rows.InsertCopy(row + 1, workingDays - 1, ws.Rows[row]);

        // Fill inserted rows with sample data.
        var random = new Random();
        for (int i = 0; i < workingDays; i++)
        {
            ExcelRow currentRow = ws.Rows[row + i];
            currentRow.Cells[1].SetValue(startDate.AddDays(i));
            currentRow.Cells[2].SetValue(random.Next(1, 12));
        }

        // Calculate formulas in worksheet.
        ws.Calculate();

        ef.Save("Template Use.xlsx");
    }
}
