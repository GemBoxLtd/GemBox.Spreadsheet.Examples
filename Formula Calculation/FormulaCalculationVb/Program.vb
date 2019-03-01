Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Formula Calculation")

        ' Some formatting.
        Dim row = worksheet.Rows(0)
        row.Style.Font.Weight = ExcelFont.BoldWeight

        Dim column = worksheet.Columns(0)
        column.SetWidth(250, LengthUnit.Pixel)
        column.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left
        column = worksheet.Columns(1)
        column.SetWidth(250, LengthUnit.Pixel)
        column.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right

        ' Use first row for column headers.
        worksheet.Cells("A1").Value = "Formula"
        worksheet.Cells("B1").Value = "Calculated value"

        ' Enter some Excel formulas as text in first column.
        worksheet.Cells("A2").Value = "=1 + 1"
        worksheet.Cells("A3").Value = "=3 * (2 - 8)"
        worksheet.Cells("A4").Value = "=3 + ABS(B3)"
        worksheet.Cells("A5").Value = "=B4 > 15"
        worksheet.Cells("A6").Value = "=IF(B5, ""Hello world"", ""World hello"")"
        worksheet.Cells("A7").Value = "=B6 & "" example"""
        worksheet.Cells("A8").Value = "=CODE(RIGHT(B7))"
        worksheet.Cells("A9").Value = "=POWER(B8, 3) * 0.45%"
        worksheet.Cells("A10").Value = "=SIGN(B9)"
        worksheet.Cells("A11").Value = "=SUM(B2:B10)"

        ' Set text from first column as second row cell's formula.
        Dim rowIndex As Integer = 0
        While worksheet.Cells(rowIndex, 0).ValueType <> CellValueType.Null
            worksheet.Cells(rowIndex, 1).Formula = worksheet.Cells(rowIndex, 0).StringValue
            rowIndex += 1
        End While

        ' GemBox.Spreadsheet supports single Excel cell calculation, ...
        worksheet.Cells("B1").Calculate()

        ' ... Excel worksheet calculation,
        worksheet.Calculate()

        ' ... and whole Excel file calculation.
        worksheet.Parent.Calculate()

        workbook.Save("Formula Calculation.xlsx")
    End Sub
End Module