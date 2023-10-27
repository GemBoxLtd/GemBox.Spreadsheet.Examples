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
        var worksheet = workbook.Worksheets.Add("InlineTextFormatting");

        worksheet.Columns[0].Width = 50 * 256;

        var cell = worksheet.Cells["A1"];
        cell.Value = "This is big and red text!";

        // Apply the size to "big and red" part of the text.
        cell.GetCharacters(8, 11).Font.Size = 400;

        // Apply the color to "red" part of the text.
        cell.GetCharacters(16, 3).Font.Color = SpreadsheetColor.FromName(ColorName.Red);

        cell = worksheet.Cells["A3"];
        cell.Value = "Formatting selected characters with GemBox.Spreadsheet component.";

        // Apply formatting on the whole cell content.
        cell.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);
        cell.Style.Font.Italic = true;
        cell.Style.WrapText = true;

        // Get characters from index 36 to the end of string,
        // e.g. the "GemBox.Spreadsheet component." part of the text.
        var characters = cell.GetCharacters(36);

        // Apply the color and underline to selected characters.
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
        characters.Font.UnderlineStyle = UnderlineStyle.Single;

        // Write selected characters.
        worksheet.Cells["A5"].Value = "Selected characters: " + characters.Text;

        workbook.Save("Inline Text Formatting.xlsx");
    }

    static void Example2()
    {
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("HtmlTextFormatting");

        worksheet.Columns[0].Width = 50 * 256;

        var htmlOptions = new HtmlLoadOptions();
        var html = "<h1 style='background:#DDEBF7'>HTML formatted text!</h1>";

        worksheet.Cells["A1"].SetValue(html, htmlOptions);

        html = @"<div style='font:11pt Calibri'>
<p>This is <span style='font-size:20pt'>big and <span style='color:red'>red</span></span> text!</p>
<p>This is <sub>subscript</sub>, <sup>superscript</sup>, <strike>strike</strike>, and <u>underline</u> text.</p>
</div>";

        worksheet.Cells["A3"].SetValue(html, htmlOptions);

        workbook.Save("Html Text Formatting.xlsx");
    }
}
