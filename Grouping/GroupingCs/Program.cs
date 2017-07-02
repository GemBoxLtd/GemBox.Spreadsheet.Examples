using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Grouping");

        ws.Cells[0].Value = "Cell grouping examples:";

        // Vertical grouping.
        ws.Cells[2, 0].Value = "GroupA Start";
        ws.Rows[2].OutlineLevel = 1;
        ws.Cells[3, 0].Value = "A";
        ws.Rows[3].OutlineLevel = 1;
        ws.Cells[4, 1].Value = "GroupB Start";
        ws.Rows[4].OutlineLevel = 2;
        ws.Cells[5, 1].Value = "B";
        ws.Rows[5].OutlineLevel = 2;
        ws.Cells[6, 1].Value = "GroupB End";
        ws.Rows[6].OutlineLevel = 2;
        ws.Cells[7, 0].Value = "GroupA End";
        ws.Rows[7].OutlineLevel = 1;
        // Put outline row buttons above groups.
        ws.ViewOptions.OutlineRowButtonsBelow = false;

        // Horizontal grouping (collapsed).
        ws.Cells["E2"].Value = "Gr.C Start";
        ws.Columns["E"].OutlineLevel = 1;
        ws.Columns["E"].Hidden = true;
        ws.Cells["F2"].Value = "C";
        ws.Columns["F"].OutlineLevel = 1;
        ws.Columns["F"].Hidden = true;
        ws.Cells["G2"].Value = "Gr.C End";
        ws.Columns["G"].OutlineLevel = 1;
        ws.Columns["G"].Hidden = true;
        ws.Columns["H"].Collapsed = true;

        ef.Save("Grouping.xlsx");
    }
}
