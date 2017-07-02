using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("TemplateUse.xlsx");

        // Get template sheet.
        ExcelWorksheet templateSheet = ef.Worksheets[0];

        // Copy template sheet.
        for (int i = 0; i < 4; i++)
            ef.Worksheets.AddCopy("Invoice " + (i + 1), templateSheet);

        // Delete template sheet.
        ef.Worksheets.Remove(0);

        DateTime startTime = DateTime.Now;

        // Go to the first Monday from today.
        while (startTime.DayOfWeek != DayOfWeek.Monday)
            startTime = startTime.AddDays(1);

        Random rnd = new Random();

        // For each sheet.
        for (int i = 0; i < 4; i++)
        {
            // Get sheet.
            ExcelWorksheet ws = ef.Worksheets[i];

            // Set some fields.
            ws.Cells["J5"].SetValue(14 + i);
            ws.Cells["J6"].SetValue(DateTime.Now);
            ws.Cells["J6"].Style.NumberFormat = "m/dd/yyyy";

            ws.Cells["D12"].Value = "ACME Corp";
            ws.Cells["D13"].Value = "240 Old Country Road, Springfield, IL";
            ws.Cells["D14"].Value = "USA";
            ws.Cells["D15"].Value = "Joe Smith";

            ws.Cells["E18"].Value = String.Format(startTime.ToShortDateString() + " until " + startTime.AddDays(11).ToShortDateString());

            for (int j = 0; j < 10; j++)
            {
                ws.Cells[21 + j, 1].SetValue(startTime); // Set date.
                ws.Cells[21 + j, 1].Style.NumberFormat = "dddd, mmmm dd, yyyy";
                ws.Cells[21 + j, 4].SetValue(rnd.Next(6, 9)); // Work hours.

                // Skip Saturday and Sunday.
                startTime = startTime.AddDays(j == 4 ? 3 : 1);
            }

            // Skip Saturday and Sunday.
            startTime = startTime.AddDays(2);

            ws.Cells["B36"].Value = "Payment via check.";
        }

        ef.Save("Sheet Copying_Deleting.xlsx");
    }
}
