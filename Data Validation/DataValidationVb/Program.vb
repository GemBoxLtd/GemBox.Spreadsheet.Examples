Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Data Validation")

        worksheet.Cells(0, 0).Value = "Data validation examples:"

        worksheet.Cells(2, 1).Value = "Decimal greater than 3.14 (on entire row 4):"
        worksheet.DataValidations.Add(New DataValidation(worksheet.Rows(3).Cells) With {
             .Type = DataValidationType.Decimal,
             .Operator = DataValidationOperator.GreaterThan,
             .Formula1 = 3.14, .InputMessageTitle = "Enter a decimal",
             .InputMessage = "Decimal should be greater than 3.14.",
             .ErrorTitle = "Invalid decimal",
             .ErrorMessage = "Value should be a decimal greater than 3.14."
        })
        worksheet.Cells.GetSubrange("A4", "J4").Value = 3.15

        worksheet.Cells(7, 1).Value = "List from B9 to B12 (on cell C8):"
        worksheet.Cells(8, 1).Value = "John"
        worksheet.Cells(9, 1).Value = "Fred"
        worksheet.Cells(10, 1).Value = "Hans"
        worksheet.Cells(11, 1).Value = "Ivan"
        worksheet.DataValidations.Add(New DataValidation(worksheet, "C8") With {
            .Type = DataValidationType.List,
            .Formula1 = "=B9:B12",
            .InputMessageTitle = "Enter a name",
            .InputMessage = "Name should be from the list: John, Fred, Hans, Ivan.",
            .ErrorStyle = DataValidationErrorStyle.Warning,
            .ErrorTitle = "Invalid name",
            .ErrorMessage = "Value should be a name from the list: John, Fred, Hans, Ivan."
        })
        worksheet.Cells("C8").Value = "John"

        worksheet.Cells(13, 1).Value = "Date between 2011-01-01 and 2011-12-31 (on cell range C14:E15):"
        worksheet.DataValidations.Add(New DataValidation(worksheet.Cells.GetSubrange("C14", "E15")) With {
            .Type = DataValidationType.Date,
            .Operator = DataValidationOperator.Between,
            .Formula1 = New DateTime(2011, 1, 1),
            .Formula2 = New DateTime(2011, 12, 31),
            .InputMessageTitle = "Enter a date",
            .InputMessage = "Date should be between 2011-01-01 and 2011-12-31.",
            .ErrorStyle = DataValidationErrorStyle.Information,
            .ErrorTitle = "Invalid date",
            .ErrorMessage = "Value should be a date between 2011-01-01 and 2011-12-31."
        })
        worksheet.Cells.GetSubrange("C14", "E15").Value = New DateTime(2011, 1, 1)

        ' Column width of 8, 55 and 15 characters.
        worksheet.Columns(0).Width = 8 * 256
        worksheet.Columns(1).Width = 55 * 256
        worksheet.Columns(2).Width = 15 * 256
        worksheet.Columns(3).Width = 15 * 256
        worksheet.Columns(4).Width = 15 * 256

        workbook.Save("Data Validation.xlsx")
    End Sub
End Module