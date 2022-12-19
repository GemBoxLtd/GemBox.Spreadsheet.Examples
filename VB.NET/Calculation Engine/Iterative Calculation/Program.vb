Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()

        ' Set calculation options.
        workbook.CalculationOptions.MaximumIterations = 10
        workbook.CalculationOptions.MaximumChange = 0.05
        workbook.CalculationOptions.EnableIterativeCalculation = True

        ' Add new worksheet.
        Dim worksheet = workbook.Worksheets.Add("Iterative Calculation")

        ' Some column formatting.
        worksheet.Columns(0).SetWidth(50, LengthUnit.Pixel)
        worksheet.Columns(1).SetWidth(100, LengthUnit.Pixel)

        ' Simple example of circular reference limited by MaximumIterations in column A.
        worksheet.Cells("A1").Formula = "=A2"
        worksheet.Cells("A2").Formula = "=A1 + 1"

        ' Simple example of circular reference limited by MaximumChange in column B.
        worksheet.Cells("B1").Value = 100000.0
        worksheet.Cells("B2").Formula = "=B3 * 0.03"
        worksheet.Cells("B3").Formula = "=B1 + B2"

        ' Calculate all cells.
        worksheet.Calculate()

        workbook.Save("Iterative Calculation.xlsx")

    End Sub
End Module