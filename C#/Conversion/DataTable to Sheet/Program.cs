using GemBox.Spreadsheet;
using System.Data;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("DataTable to Sheet");

        var dataTable = new DataTable();

        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("FirstName", typeof(string));
        dataTable.Columns.Add("LastName", typeof(string));

        dataTable.Rows.Add(new object[] { 100, "John", "Doe" });
        dataTable.Rows.Add(new object[] { 101, "Fred", "Nurk" });
        dataTable.Rows.Add(new object[] { 103, "Hans", "Meier" });
        dataTable.Rows.Add(new object[] { 104, "Ivan", "Horvat" });
        dataTable.Rows.Add(new object[] { 105, "Jean", "Dupont" });
        dataTable.Rows.Add(new object[] { 106, "Mario", "Rossi" });

        worksheet.Cells[0, 0].Value = "DataTable insert example:";

        // Insert DataTable to an Excel worksheet.
        worksheet.InsertDataTable(dataTable,
            new InsertDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 2
            });

        workbook.Save("DataTable to Sheet.xlsx");
    }
}
