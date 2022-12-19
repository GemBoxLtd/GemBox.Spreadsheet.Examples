using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formula Utility Methods");

        // Fill first column with values.
        for (int i = 0; i < 10; i++)
            worksheet.Cells[i, 0].Value = i + 1;

        // Cell B1 has formula '=A1*2', B2 '=A2*2', etc.
        for (int i = 0; i < 10; i++)
            worksheet.Cells[i, 1].Formula = String.Format("={0}*2", CellRange.RowColumnToPosition(i, 0));

        // Cell C1 has formula '=SUM(A1:B1)', C2 '=SUM(A2:B2)', etc.
        for (int i = 0; i < 10; i++)
            worksheet.Cells[i, 2].Formula = String.Format("=SUM(A{0}:B{0})", ExcelRowCollection.RowIndexToName(i));

        // Cell A12 contains sum of all values from the first row.
        worksheet.Cells["A12"].Formula = String.Format("=SUM(A1:{0}1)", ExcelColumnCollection.ColumnIndexToName(worksheet.Rows[0].AllocatedCells.Count - 1));

        workbook.Save("Formula Utility Methods.xlsx");
    }
}