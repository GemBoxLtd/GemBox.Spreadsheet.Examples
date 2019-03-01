using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Hyperlinks");

        worksheet.Cells["A1"].Value = "Hyperlink examples:";

        var cell = worksheet.Cells["B3"];
        cell.Value = "GemBoxSoftware";
        cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
        cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);
        cell.Hyperlink.Location = "https://www.gemboxsoftware.com";
        cell.Hyperlink.IsExternal = true;

        cell = worksheet.Cells["B5"];
        cell.Value = "Jump";
        cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
        cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);
        cell.Hyperlink.ToolTip = "This is tool tip! This hyperlink jumps to E1.";
        cell.Hyperlink.Location = worksheet.Name + "!E1";

        worksheet.Cells["E1"].Value = "Destination";

        cell = worksheet.Cells["B8"];
        cell.Formula = "=HYPERLINK(\"https://www.gemboxsoftware.com/spreadsheet/examples/excel-cell-hyperlinks/207\", \"Example of HYPERLINK formula\")";
        cell.Style.Font.UnderlineStyle = UnderlineStyle.Single;
        cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);

        workbook.Save("Hyperlinks.xlsx");
    }
}
