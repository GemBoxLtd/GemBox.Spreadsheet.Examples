Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile

        ' Set calculation options.
        ef.CalculationOptions.MaximumIterations = 10
        ef.CalculationOptions.MaximumChange = 0.05
        ef.CalculationOptions.EnableIterativeCalculation = True

        ' Add new worksheet.
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Iterative Calculation")

        ' Some column formatting.
        ws.Columns(0).SetWidth(50, LengthUnit.Pixel)
        ws.Columns(1).SetWidth(100, LengthUnit.Pixel)

        ' Simple example of circular reference limited by MaximumIterations in column A.
        ws.Cells("A1").Formula = "=A2"
        ws.Cells("A2").Formula = "=A1 + 1"

        ' Simple example of circular reference limited by MaximumChange in column B.
        ws.Cells("B1").Value = 100000.0
        ws.Cells("B2").Formula = "=B3 * 0.03"
        ws.Cells("B3").Formula = "=B1 + B2"

        ' Calculate all cells.
        ws.Calculate()
        ef.Save("Iterative Calculation.xlsx")

    End Sub

End Module