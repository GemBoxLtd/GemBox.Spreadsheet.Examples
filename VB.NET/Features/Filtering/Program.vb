Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SampleData.xlsx")
        Dim worksheet = workbook.Worksheets("Data")
        Dim rowCount As Integer = worksheet.Rows.Count

        ' Specify range which will be filtered.
        Dim filterRange = worksheet.Cells.GetSubrangeAbsolute(0, 0, rowCount, 4)

        ' Show only rows which satisfy following conditions:
        ' - 'Departments' value is either "Legal" or "Marketing" or "Finance" and
        ' - 'Names' value contains word "Fred" and
        ' - 'Salaries' value is in the top 20 percent of all 'Salaries' values.
        ' Shown rows are then sorted by 'Salaries' values in the descending order.
        filterRange.Filter() _
            .ByValues(0, "Legal", "Marketing", "Finance") _
            .ByCustom(1, FilterOperator.Equal, "*Fred*") _
            .ByTop10(3, True, True, 20) _
            .SortBy(3, True) _
            .Apply()

        workbook.Save("Filtering.xlsx")

    End Sub

End Module
