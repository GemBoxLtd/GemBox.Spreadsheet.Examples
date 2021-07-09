using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Comments");

        // Add hidden comments (hover over an indicator to view it).
        ExcelCell cell = worksheet.Cells["B2"];
        cell.Value = "Hidden comment";
        ExcelComment comment = cell.Comment;
        comment.Text = "Comment with hidden text.";

        comment = worksheet.Cells["B4"].Comment;
        comment.Text = "Another comment with hidden text.";

        // Add visible comments.
        cell = worksheet.Cells["B6"];
        cell.Value = "Visible comment";
        comment = cell.Comment;
        comment.Text = "Comment with specified position and size.";
        comment.IsVisible = true;
        comment.TopLeftCell = new AnchorCell(worksheet.Cells["D5"], true);
        comment.BottomRightCell = new AnchorCell(worksheet.Cells["E12"], false);

        comment = worksheet.Cells["B8"].Comment;
        comment.Text = "Comment with specified start position.";
        comment.IsVisible = true;
        comment.TopLeftCell = new AnchorCell(worksheet.Columns["A"], worksheet.Rows["10"], 20, 10, LengthUnit.Pixel);

        // Add visible comment with formatted individual characters.
        comment = worksheet.Cells["F3"].Comment;
        comment.Text = "Comment with rich formatted text.\nComment is:\n a) multiline,\n b) large,\n c) visible, \n d) formatted, and \n e) autofitted.";
        comment.IsVisible = true;
        var characters = comment.GetCharacters(0, 33);
        characters.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
        characters.Font.Weight = ExcelFont.BoldWeight;
        characters.Font.Size = 300;
        comment.GetCharacters(13, 4).Font.Color = SpreadsheetColor.FromName(ColorName.Blue);
        comment.AutoFit();

        // Read and update comment.
        cell = worksheet.Cells["B8"];
        if (cell.Comment.Exists)
        {
            cell.Comment.Text = cell.Comment.Text.Replace(".", " and modified text.");
            cell.Value = "Updated comment.";
        }

        // Delete comment.
        cell = worksheet.Cells["B4"];
        cell.Comment = null;
        cell.Value = "Deleted comment.";

        workbook.Save("Cell Comments.xlsx");
    }
}