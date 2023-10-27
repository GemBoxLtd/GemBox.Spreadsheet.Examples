using System.Linq;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        // Get the cell range.
        var range = worksheet.Cells.GetSubrange("B2:E5");

        // Merge cells in the current range.
        range.Merged = true;

        // Set the value of the merged range.
        range.Value = "Merged";

        // Set the style of the merged range.
        range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;

        // Set the style of the merged range using a cell within.
        worksheet.Cells["C3"].Style.Borders
            .SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Double);

        workbook.Save("Merged Cells.xlsx");
    }

    static void Example2()
    {
        var workbook = ExcelFile.Load("Merged Cells.xlsx");
        var worksheet = workbook.Worksheets[0];

        // Get the first merged range.
        var mergedRange = worksheet.Rows
            .SelectMany(row => row.AllocatedCells)
            .Select(cell => cell.MergedRange)
            .FirstOrDefault(range => range != null);

        if (mergedRange != null)
        {
            // Important, you cannot unmerge the ExcelCell.MergedRange property.
            // In other words, the following is not allowed:  mergedRange.Merged = false;

            // Instead, you need to retrieve the same CellRange from the ExcelWorksheet and then unmerge it.
            // This kind of implementation was chosen for performance reasons.
            worksheet.Cells.GetSubrange(mergedRange.Name).Merged = false;

            worksheet.Cells[mergedRange.StartPosition].Value = "Unmerged";
        }

        workbook.Save("Unmerged Cells.xlsx");
    }
}