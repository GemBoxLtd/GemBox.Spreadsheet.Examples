using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Comments");

        ws.Cells[0, 0].Value = "Comment examples:";

        ws.Cells[2, 1].Comment.Text = "Empty cell.";

        ws.Cells[4, 1].Value = 5;
        ws.Cells[4, 1].Comment.Text = "Cell with a number.";

        ws.Cells["B7"].Value = "Cell B7";

        ExcelComment comment = ws.Cells["B7"].Comment;
        comment.Text = "Some formatted text.\nComment is:\na) multiline,\nb) large,\nc) visible, and \nd) formatted.";
        comment.IsVisible = true;
        comment.TopLeftCell = new AnchorCell(ws.Columns[3], ws.Rows[4], true);
        comment.BottomRightCell = new AnchorCell(ws.Columns[5], ws.Rows[10], false);

        // Get first 20 characters of a string
        FormattedCharacterRange characters = comment.GetCharacters(0, 20);

        // Apply color, italic and size to selected characters
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
        characters.Font.Italic = true;
        characters.Font.Size = 300;

        // Apply color to 'formatted' part of text
        comment.GetCharacters(5, 9).Font.Color = SpreadsheetColor.FromName(ColorName.Red);

        ef.Save("Comments.xlsx");
    }
}
