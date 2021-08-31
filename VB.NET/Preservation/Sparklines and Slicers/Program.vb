Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Excel file with preservation feature enabled.
        Dim loadOptions As New XlsxLoadOptions() With {.PreserveUnsupportedFeatures = True}
        Dim workbook = ExcelFile.Load("SparklinesAndSlicers.xlsx", loadOptions)

        ' Modify all values in column C, set them to some random value.
        Dim readEnumerator = workbook.Worksheets(0).Columns("C").Cells.GetReadEnumerator()
        Dim random As New Random()
        While readEnumerator.MoveNext()

            Dim cell = readEnumerator.Current
            If cell.ValueType = CellValueType.Int Then cell.SetValue(random.Next(-10, 10))

        End While

        workbook.Save("Preserved Output.xlsx")

    End Sub
End Module