Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("Excel 2010.xlsx")

        ' Modify all values in column C. Set them to some random value between -10 and 10.
        Dim readEnumerator = ef.Worksheets(0).Columns("C").Cells.GetReadEnumerator()

        Dim rnd = New Random()
        While readEnumerator.MoveNext()

            Dim cell = readEnumerator.Current
            If cell.ValueType = CellValueType.Int Then cell.SetValue(rnd.Next(-10, 10))

        End While

        ef.Save("Excel 2010_2013 Features.xlsx")

    End Sub

End Module