Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Filtering")

        Dim rowCount As Integer = 149

        ' Specify sheet formatting.
        worksheet.Rows(0).Style.Font.Weight = ExcelFont.BoldWeight
        worksheet.Columns(0).SetWidth(3, LengthUnit.Centimeter)
        worksheet.Columns(1).SetWidth(3, LengthUnit.Centimeter)
        worksheet.Columns(2).SetWidth(3, LengthUnit.Centimeter)
        worksheet.Columns(2).Style.NumberFormat = "[$$-409]#,##0.00"
        worksheet.Columns(3).SetWidth(3, LengthUnit.Centimeter)
        worksheet.Columns(3).Style.NumberFormat = "yyyy-mm-dd"

        Dim cells = worksheet.Cells

        ' Specify header row.
        cells(0, 0).Value = "Departments"
        cells(0, 1).Value = "Names"
        cells(0, 2).Value = "Salaries"
        cells(0, 3).Value = "Deadlines"

        ' Insert random data to sheet.
        Dim random As New Random()
        Dim departments = New String() {"Legal", "Marketing", "Finance", "Planning", "Purchasing"}
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        For i As Integer = 0 To rowCount - 1
            cells(i + 1, 0).Value = departments(random.Next(departments.Length))
            cells(i + 1, 1).Value = names(random.Next(names.Length)) + " "c + (i + 1).ToString()
            cells(i + 1, 2).SetValue(random.Next(10, 101) * 100)
            cells(i + 1, 3).SetValue(DateTime.Now.AddDays(random.Next(-1, 2)))
        Next

        ' Specify range which will be filtered.
        Dim filterRange = worksheet.Cells.GetSubrangeAbsolute(0, 0, rowCount, 3)

        ' Show only rows which satisfy following conditions:
        ' - 'Departments' value is either "Legal" or "Marketing" or "Finance" and
        ' - 'Names' value contains letter 'e' and
        ' - 'Salaries' value is in the top 20 percent of all 'Salaries' values and
        ' - 'Deadlines' value is today's date.
        ' Shown rows are then sorted by 'Salaries' values in the descending order.
        filterRange.Filter() _
            .ByValues(0, "Legal", "Marketing", "Finance") _
            .ByCustom(1, FilterOperator.Equal, "*e*") _
            .ByTop10(2, True, True, 20) _
            .ByDynamic(3, DynamicFilterType.Today) _
            .SortBy(2, True) _
            .Apply()

        workbook.Save("Filtering.xlsx")
    End Sub
End Module