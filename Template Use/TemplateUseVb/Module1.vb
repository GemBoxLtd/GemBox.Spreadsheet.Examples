Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("Template.xlsx")

        Dim workingDays As Integer = 8

        Dim startDate As DateTime = DateTime.Now.AddDays(-workingDays)
        Dim endDate As DateTime = DateTime.Now

        Dim ws As ExcelWorksheet = ef.Worksheets(0)

        ' Find cells with placeholder text and set their values.
        Dim row As Integer, column As Integer
        If ws.Cells.FindText("[Company Name]", True, True, row, column) Then
            ws.Cells(row, column).Value = "ACME Corp"
        End If
        If ws.Cells.FindText("[Company Address]", True, True, row, column) Then
            ws.Cells(row, column).Value = "240 Old Country Road, Springfield, IL"
        End If
        If ws.Cells.FindText("[Start Date]", True, True, row, column) Then
            ws.Cells(row, column).Value = startDate
        End If
        If ws.Cells.FindText("[End Date]", True, True, row, column) Then
            ws.Cells(row, column).Value = endDate
        End If

        ' Copy template row.
        row = 17
        ws.Rows.InsertCopy(row + 1, workingDays - 1, ws.Rows(row))

        ' Fill inserted rows with sample data.
        Dim random As New Random()
        For i As Integer = 0 To workingDays - 1
            Dim currentRow As ExcelRow = ws.Rows(row + i)
            currentRow.Cells(1).SetValue(startDate.AddDays(i))
            currentRow.Cells(2).SetValue(random.Next(1, 12))
        Next

        ' Calculate formulas in worksheet.
        ws.Calculate()

        ef.Save("Template Use.xlsx")

    End Sub

End Module