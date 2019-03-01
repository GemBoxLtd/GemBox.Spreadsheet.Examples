using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("TemplateUse.xlsx");

        // Get template sheet.
        var templateSheet = workbook.Worksheets[0];

        // Copy template sheet.
        for (int i = 0; i < 4; i++)
            workbook.Worksheets.AddCopy("Invoice " + (i + 1), templateSheet);

        // Delete template sheet.
        workbook.Worksheets.Remove(0);

        var startTime = DateTime.Now;

        // Go to the first Monday from today.
        while (startTime.DayOfWeek != DayOfWeek.Monday)
            startTime = startTime.AddDays(1);

        var random = new Random();

        // For each sheet.
        for (int i = 0; i < 4; i++)
        {
            // Get sheet.
            var worksheet = workbook.Worksheets[i];

            // Set some fields.
            worksheet.Cells["J5"].SetValue(14 + i);
            worksheet.Cells["J6"].SetValue(DateTime.Now);
            worksheet.Cells["J6"].Style.NumberFormat = "m/dd/yyyy";

            worksheet.Cells["D12"].Value = "ACME Corp";
            worksheet.Cells["D13"].Value = "240 Old Country Road, Springfield, IL";
            worksheet.Cells["D14"].Value = "USA";
            worksheet.Cells["D15"].Value = "Joe Smith";

            worksheet.Cells["E18"].Value = String.Format(startTime.ToShortDateString() + " until " + startTime.AddDays(11).ToShortDateString());

            for (int j = 0; j < 10; j++)
            {
                worksheet.Cells[21 + j, 1].SetValue(startTime); // Set date.
                worksheet.Cells[21 + j, 1].Style.NumberFormat = "dddd, mmmm dd, yyyy";
                worksheet.Cells[21 + j, 4].SetValue(random.Next(6, 9)); // Work hours.

                // Skip Saturday and Sunday.
                startTime = startTime.AddDays(j == 4 ? 3 : 1);
            }

            // Skip Saturday and Sunday.
            startTime = startTime.AddDays(2);

            worksheet.Cells["B36"].Value = "Payment via check.";
        }

        workbook.Save("Sheet Copying_Deleting.xlsx");
    }
}