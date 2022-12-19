using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        int numberOfItems = 10;
        var startDate = DateTime.Today.AddDays(-numberOfItems);
        var endDate = DateTime.Today;

        // Load an Excel template.
        var workbook = ExcelFile.Load("Template.xlsx");

        // Get template sheet.
        var worksheet = workbook.Worksheets[0];

        // Find cells with placeholder text and set their values.
        int row, column;
        if (worksheet.Cells.FindText("[Company Name]", out row, out column))
            worksheet.Cells[row, column].Value = "ACME Corp";
        if (worksheet.Cells.FindText("[Company Address]", out row, out column))
            worksheet.Cells[row, column].Value = "240 Old Country Road, Springfield, IL";
        if (worksheet.Cells.FindText("[Start Date]", out row, out column))
            worksheet.Cells[row, column].Value = startDate;
        if (worksheet.Cells.FindText("[End Date]", out row, out column))
            worksheet.Cells[row, column].Value = endDate;

        // Copy template row.
        row = 17;
        worksheet.Rows.InsertCopy(row + 1, numberOfItems - 1, worksheet.Rows[row]);

        // Fill copied rows with sample data.
        var random = new Random();
        for (int i = 0; i < numberOfItems; i++)
        {
            var currentRow = worksheet.Rows[row + i];
            currentRow.Cells[1].SetValue(startDate.AddDays(i));
            currentRow.Cells[2].SetValue(random.Next(1, 12));
        }

        // Calculate formulas in a sheet.
        worksheet.Calculate();

        // Save the modified Excel template to output file.
        workbook.Save("Output.xlsx");
    }
}