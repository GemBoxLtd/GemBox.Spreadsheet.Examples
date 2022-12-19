using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("RightToLeft.xlsx");
        var worksheet = workbook.Worksheets[0];

        // Show columns from the right side of the page.
        worksheet.ViewOptions.ShowColumnsFromRightToLeft = true;


        worksheet.Cells["A8"].Value = "200 جديدة";
        // Set the reading order of the cell as right-to-left.
        worksheet.Cells["A8"].Style.TextDirection = TextDirection.RightToLeft;

        workbook.Save("RightToLeft.pdf");
    }
}