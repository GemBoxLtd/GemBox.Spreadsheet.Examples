Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("TemplateUse.xlsx")

        ' Add Sheet.
        Dim worksheet As ExcelWorksheet = workbook.Worksheets.InsertEmpty(0, "Document Properties")
        workbook.Worksheets.ActiveWorksheet = worksheet

        Dim rowIndex As Integer = 0
        ' Read Built-in Document Properties.
        worksheet.Cells(rowIndex, 0).Value = "Built-in document properties"
        rowIndex = rowIndex + 1

        worksheet.Cells(rowIndex, 0).Value = "Property"
        worksheet.Cells(rowIndex, 1).Value = "Value"
        rowIndex = rowIndex + 1

        For Each keyValue In workbook.DocumentProperties.BuiltIn

            worksheet.Cells(rowIndex, 0).Value = keyValue.Key.ToString()
            worksheet.Cells(rowIndex, 1).Value = keyValue.Value
            rowIndex = rowIndex + 1
        Next

        ' Read Custom Document Properties.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Custom Document Properties"

        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Property"
        worksheet.Cells(rowIndex, 1).Value = "Value"
        rowIndex = rowIndex + 1

        For Each keyValue In workbook.DocumentProperties.Custom

            worksheet.Cells(rowIndex, 0).Value = keyValue.Key
            worksheet.Cells(rowIndex, 1).Value = keyValue.Value.ToString()
            rowIndex = rowIndex + 1
        Next

        ' Write/Modify Document Properties.
        workbook.DocumentProperties.BuiltIn(BuiltInDocumentProperties.Author) = "John Doe"
        workbook.DocumentProperties.BuiltIn(BuiltInDocumentProperties.Title) = "Generated title"

        worksheet.Columns(0).SetWidth(192, LengthUnit.Pixel)
        worksheet.Columns(1).SetWidth(217, LengthUnit.Pixel)

        workbook.Save("Document Properties.xlsx")
    End Sub
End Module