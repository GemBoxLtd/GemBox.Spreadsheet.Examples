using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, remove this FreeLimitReached event handler.
        SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        var workbook = ExcelFile.Load("SampleData.xlsx");
        var worksheet = workbook.Worksheets["Data"];
        int rowCount = worksheet.Rows.Count;

        // Specify range which will be filtered.
        var filterRange = worksheet.Cells.GetSubrangeAbsolute(0, 0, rowCount, 4);

        // Show only rows which satisfy following conditions:
        // - 'Departments' value is either "Legal" or "Marketing" or "Finance" and
        // - 'Names' value contains letter 'e' and
        // - 'Salaries' value is in the top 20 percent of all 'Salaries' values and
        // - 'Deadlines' value is today's date.
        // Shown rows are then sorted by 'Salaries' values in the descending order.
        filterRange.Filter()
            .ByValues(0, "Legal", "Marketing", "Finance")
            .ByCustom(1, FilterOperator.Equal, "*e*")
            .ByTop10(3, true, true, 20)
            .ByDynamic(4, DynamicFilterType.Today)
            .SortBy(3, true)
            .Apply();

        workbook.Save("Filtering.xlsx");
    }
}