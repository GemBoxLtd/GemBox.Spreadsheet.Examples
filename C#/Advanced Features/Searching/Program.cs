using System;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SimpleTemplate.xlsx");

        var searchText = "Apollo 13";

        var worksheet = workbook.Worksheets[0];

        var sb = new StringBuilder();

        int objectRow, objectColumn;
        worksheet.Cells.FindText(searchText, false, false, out objectRow, out objectColumn);

        if (objectRow == -1 || objectColumn == -1)
            sb.AppendLine("Can't find text.");
        else
        {
            sb.AppendLine(searchText + " was launched on " + worksheet.Cells[objectRow, 2].Value.ToString() + ".");

            var nationality = worksheet.Cells[objectRow, 1].Value as string;
            if (nationality != null)
            {
                var nationalityText = nationality.Trim().ToLowerInvariant();

                int nationalityCounter = 0;

                var enumerator = worksheet.Columns[1].Cells.GetReadEnumerator();
                while (enumerator.MoveNext())
                {
                    var cell = enumerator.Current;
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