using System;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = ExcelFile.Load("SimpleTemplate.xlsx");

        string searchText = "Apollo 13";

        var ws = ef.Worksheets[0];

        StringBuilder sb = new StringBuilder();

        int objectRow, objectColumn;
        ws.Cells.FindText(searchText, false, false, out objectRow, out objectColumn);

        if (objectRow == -1 || objectColumn == -1)
            sb.AppendLine("Can't find text.");
        else
        {
            sb.AppendLine(searchText + " was launched on " + ws.Cells[objectRow, 2].Value + ".");

            string nationality = ws.Cells[objectRow, 1].Value as string;
            if (nationality != null)
            {
                string nationalityText = nationality.Trim().ToLowerInvariant();

                int nationalityCounter = 0;

                CellRangeEnumerator enumerator = ws.Columns[1].Cells.GetReadEnumerator();
                while (enumerator.MoveNext())
                {
                    ExcelCell cell = enumerator.Current;
                    var cellValue = cell.Value as string;
                    if (cellValue != null && cellValue.Trim().ToLowerInvariant() == nationalityText)
                        nationalityCounter++;
                }

                sb.AppendFormat("There are {0} entires for {1}.", nationalityCounter, nationality);
            }
        }

        Console.WriteLine(sb.ToString());
    }
}
