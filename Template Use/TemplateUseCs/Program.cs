using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("Template.xlsx");

        int workingDays = 8;

        var startDate = DateTime.Now.AddDays(-workingDays);
        var endDate = DateTime.Now;

        var worksheet = workbook.Worksheets[0];

        // Find cells with placeholder text and set their values.
        int row, column;
        if (worksheet.Cells.FindText("[Company Name]", true, true, out row, out column))
            worksheet.Cells[row, column].Value = "ACME Corp";
        if (worksheet.Cells.FindText("[Company Address]", true, true, out row, out column))
            worksheet.Cells[row, column].Value = "240 Old Country Road, Springfield, IL";
        if (worksheet.Cells.FindText("[Start Date]", true, true, out row, out column))
            worksheet.Cells[row, column].Value = startDate;
        if (worksheet.Cells.FindText("[End Date]", true, true, out row, out column))
            worksheet.Cells[row, column].Value = endDate;

        // Copy template row.
        row = 17;
        worksheet.Rows.InsertCopy(row + 1, workingDays - 1, worksheet.Rows[row]);

        // Fill inserted rows with sample data.
        var random = new Random();
        for (int i = 0; i < workingDays; i++)
        {
            var currentRow = worksheet.Rows[row + i];
            currentRow.Cells[1].SetValue(startDate.AddDays(i));
            currentRow.Cells[2].SetValue(random.Next(1, 12));
        }

        // Calculate formulas in worksheet.
        worksheet.Calculate();

        workbook.Save("Template Use.xlsx");
    }
}