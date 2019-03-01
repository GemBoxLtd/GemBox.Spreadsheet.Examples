using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Comments");

        worksheet.Cells[0, 0].Value = "Comment examples:";

        worksheet.Cells[2, 1].Comment.Text = "Empty cell.";

        worksheet.Cells[4, 1].Value = 5;
        worksheet.Cells[4, 1].Comment.Text = "Cell with a number.";

        worksheet.Cells["B7"].Value = "Cell B7";

        var comment = worksheet.Cells["B7"].Comment;
        comment.Text = "Some formatted text.\nComment is:\na) multiline,\nb) large,\nc) visible, and \nd) formatted.";
        comment.IsVisible = true;
        comment.TopLeftCell = new AnchorCell(worksheet.Columns[3], worksheet.Rows[4], true);
        comment.BottomRightCell = new AnchorCell(worksheet.Columns[5], worksheet.Rows[10], false);

        // Get first 20 characters of a string.
        var characters = comment.GetCharacters(0, 20);

        // Apply color, italic and size to selected characters.
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
        characters.Font.Italic = true;
        characters.Font.Size = 300;

        // Apply color to 'formatted' part of text.
        comment.GetCharacters(5, 9).Font.Color = SpreadsheetColor.FromName(ColorName.Red);

        workbook.Save("Comments.xlsx");
    }
}