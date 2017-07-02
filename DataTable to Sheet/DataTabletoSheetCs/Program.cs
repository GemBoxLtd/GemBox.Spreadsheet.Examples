using System.Data;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");


        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("DataTable to Sheet");

        DataTable dt = new DataTable();

        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("FirstName", typeof(string));
        dt.Columns.Add("LastName", typeof(string));

        dt.Rows.Add(new object[] { 100, "John", "Doe" });
        dt.Rows.Add(new object[] { 101, "Fred", "Nurk" });
        dt.Rows.Add(new object[] { 103, "Hans", "Meier" });
        dt.Rows.Add(new object[] { 104, "Ivan", "Horvat" });
        dt.Rows.Add(new object[] { 105, "Jean", "Dupont" });
        dt.Rows.Add(new object[] { 106, "Mario", "Rossi" });

        ws.Cells[0, 0].Value = "DataTable insert example:";

        // Insert DataTable into an Excel worksheet.
        ws.InsertDataTable(dt,
            new InsertDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 2
            });

        ef.Save("DataTable to Sheet.xlsx");
    }
}
