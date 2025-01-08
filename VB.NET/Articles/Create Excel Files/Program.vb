Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Charts
Imports GemBox.Spreadsheet.PivotTables
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create an Excel file.
        Dim workbook As New ExcelFile()

        ' Add a new worksheet to the Excel file.
        Dim worksheet = workbook.Worksheets.Add("New worksheet")

        ' Set the value of the cell "A1".
        worksheet.Cells("A1").Value = "Hello world!"

        ' Save the Excel file to a file format of your choice.
        workbook.Save("Create.xlsx")

        ' Set the value of the cell "A1".
        worksheet.Cells("A1").Value = "Some important information"

        ' Apply bold formatting to the cell
        worksheet.Cells("A1").Style.Font.Weight = ExcelFont.BoldWeight

        ' Set the top row background color
        worksheet.Rows(0).Style.FillPattern.SetSolid(
            SpreadsheetColor.FromName(ColorName.LightBlue))

        ' Set the top row borders to thin black
        worksheet.Rows(0).Style.Borders.SetBorders(
            MultipleBorders.All,
            SpreadsheetColor.FromName(ColorName.Black),
            LineStyle.Thin)

        ' Automatically adjust column width based on content
        worksheet.Columns(0).AutoFit()

        ' Add a header with report title and date
        Dim header = worksheet.HeadersFooters.DefaultPage.Header
        header.CenterSection.Append("Report Title")
        header.RightSection.Append(HeaderFooterFieldType.Date)

        ' Add a footer with page number
        Dim footer = worksheet.HeadersFooters.DefaultPage.Footer
        footer.RightSection.Append(HeaderFooterFieldType.PageNumber)

        ' Add new worksheet for sales data
        Dim dataSheet = workbook.Worksheets.Add("Data")

        ' Sample dat
        dataSheet.Cells("A1").Value = "Month"
        dataSheet.Cells("A2").Value = "January"
        dataSheet.Cells("A3").Value = "February"
        dataSheet.Cells("A4").Value = "March"
        dataSheet.Cells("A5").Value = "April"
        dataSheet.Cells("A6").Value = "May"
        dataSheet.Cells("A7").Value = "June"
        dataSheet.Cells("A8").Value = "July"
        dataSheet.Cells("A9").Value = "August"
        dataSheet.Cells("A10").Value = "September"
        dataSheet.Cells("A11").Value = "October"
        dataSheet.Cells("A12").Value = "November"
        dataSheet.Cells("A13").Value = "December"

        dataSheet.Cells("B1").Value = "Sales"
        dataSheet.Cells("B2").Value = 5000
        dataSheet.Cells("B3").Value = 7000
        dataSheet.Cells("B4").Value = 8000
        dataSheet.Cells("B5").Value = 4500
        dataSheet.Cells("B6").Value = 3000
        dataSheet.Cells("B7").Value = 9000
        dataSheet.Cells("B8").Value = 5500
        dataSheet.Cells("B9").Value = 6000
        dataSheet.Cells("B10").Value = 9500
        dataSheet.Cells("B11").Value = 8500
        dataSheet.Cells("B12").Value = 4000
        dataSheet.Cells("B13").Value = 5000

        ' Apply conditional formatting to sales data
        Dim condition = dataSheet.ConditionalFormatting.AddTopOrBottomRanked("B2:B13", False, 3)
        condition.Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.LightGreen)
        condition.Style.Font.Weight = ExcelFont.BoldWeight

        ' Create a column chart and select data for it.
        Dim chart = dataSheet.Charts.Add(ChartType.Column, "D2", "M25")
        chart.SelectData(dataSheet.Cells.GetSubrange("A1:B13"), True)

        ' Create a new sheet for the pivot table.
        Dim pivotSheet = workbook.Worksheets.Add("PivotTable")

        ' Specify header row.
        pivotSheet.Cells("A1").Value = "Product"
        pivotSheet.Cells("B1").Value = "Region"
        pivotSheet.Cells("C1").Value = "Sales"

        ' Insert random data to the sheet.
        Dim random As New Random()
        Dim products = New String() {"Product A", "Product B", "Product C"}
        Dim regions = New String() {"North", "South", "East", "West", "Northeast"}
        Dim rowIndex = 0
        For Each product In products
            For Each region In regions
                rowIndex += 1
                pivotSheet.Cells(rowIndex, 0).Value = product
                pivotSheet.Cells(rowIndex, 1).Value = region
                pivotSheet.Cells(rowIndex, 2).Value = random.Next(2000, 9800)
            Next
        Next

        ' Create pivot cache from cell range "PivotTable!A1:C16".
        Dim cache = workbook.PivotCaches.AddWorksheetSource("PivotTable!A1:C16")

        ' Create pivot table "Product Sales" using the specified pivot cache and add it to the worksheet at the cell location 'E1'.
        Dim pivotTable = pivotSheet.PivotTables.Add(cache, "Product Sales", "E1")

        ' Aggregate 'Sales' values with sum.
        Dim field = pivotTable.DataFields.Add("Sales")
        field.Function = PivotFieldCalculationType.Sum
        field.Name = "Total Sales"
        field.NumberFormat = "[$$-409]#,##0.00"

        ' Group rows into 'Products'.
        pivotTable.RowFields.Add("Product")

        ' Group columns into 'Regions'.
        pivotTable.ColumnFields.Add("Region")

        ' Specify the string to be displayed in row and column header.
        pivotTable.RowHeaderCaption = "Products"
        pivotTable.ColumnHeaderCaption = "Regions"

        ' Create a new sheet for data validation.
        Dim dataValidationSheet = workbook.Worksheets.Add("DataValidation")

        ' Add data validation for cell 'C2' to only accept values from 'B2:B5' range.
        dataValidationSheet.Cells("B1").Value = "List from B2 to B5 (on cell C2):"
        dataValidationSheet.Cells("B2").Value = "John"
        dataValidationSheet.Cells("B3").Value = "Fred"
        dataValidationSheet.Cells("B4").Value = "Hans"
        dataValidationSheet.Cells("B5").Value = "Ivan"
        dataValidationSheet.DataValidations.Add(New DataValidation(worksheet, "C2") With
        {
            .Type = DataValidationType.List,
            .Formula1 = "=B2:B5",
            .InputMessageTitle = "Enter a name",
            .InputMessage = "Name must be from the list: John, Fred, Hans, Ivan.",
            .ErrorStyle = DataValidationErrorStyle.Stop,
            .ErrorTitle = "Invalid name",
            .ErrorMessage = "Value must be a name from the list: John, Fred, Hans, Ivan."
        })

        ' Save report to file
        workbook.Save("Report.xlsx")

        ' Convert the report to another format
        workbook.Save("Report.xlsb")

    End Sub

End Module