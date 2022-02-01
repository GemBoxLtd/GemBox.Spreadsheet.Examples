using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("Template.xlsx");

        // Get template sheet.
        var templateSheet = workbook.Worksheets[0];

        // Copy template sheet.
        for (int i = 0; i < 4; i++)
            workbook.Worksheets.AddCopy("Invoice " + (i + 1), templateSheet);

        // Delete template sheet.
        workbook.Worksheets.Remove(0);

        var random = new Random();

        // For each sheet.
        for (int i = 0; i < 4; i++)
        {
            // Get sheet.
            var worksheet = workbook.Worksheets[i];

            // Write sheet's cells.
            worksheet.Cells["C6"].Value = "ACME Corp";
            worksheet.Cells["C7"].Value = "240 Old Country Road, Springfield, IL";

            DateTime startDate = DateTime.Today;
            int itemsCount = random.Next(5, 20);
            worksheet.Cells["C11"].SetValue(startDate);
            worksheet.Cells["C12"].SetValue(startDate.AddDays(itemsCount - 1));

            // Copy template row.
            int row = 17;
            worksheet.Rows.InsertCopy(row + 1, itemsCount - 1, worksheet.Rows[row]);

            // Write row's cells.
            for (int j = 0; j < itemsCount; j++)
            {
                var currentRow = worksheet.Rows[row + j];
                currentRow.Cells[1].SetValue(startDate.AddDays(j));
                currentRow.Cells[2].SetValue(random.Next(6, 9));
            }
        }

        workbook.Save("Sheet Copying_Deleting.xlsx");
    }
}