Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Template.xlsx")

        Dim workingDays As Integer = 8

        Dim startDate = DateTime.Now.AddDays(-workingDays)
        Dim endDate = DateTime.Now

        Dim worksheet = workbook.Worksheets(0)

        ' Find cells with placeholder text and set their values.
        Dim row As Integer, column As Integer
        If worksheet.Cells.FindText("[Company Name]", True, True, row, column) Then
            worksheet.Cells(row, column).Value = "ACME Corp"
        End If
        If worksheet.Cells.FindText("[Company Address]", True, True, row, column) Then
            worksheet.Cells(row, column).Value = "240 Old Country Road, Springfield, IL"
        End If
        If worksheet.Cells.FindText("[Start Date]", True, True, row, column) Then
            worksheet.Cells(row, column).Value = startDate
        End If
        If worksheet.Cells.FindText("[End Date]", True, True, row, column) Then
            worksheet.Cells(row, column).Value = endDate
        End If

        ' Copy template row.
        row = 17
        worksheet.Rows.InsertCopy(row + 1, workingDays - 1, worksheet.Rows(row))

        ' Fill inserted rows with sample data.
        Dim random As New Random()
        For i As Integer = 0 To workingDays - 1
            Dim currentRow = worksheet.Rows(row + i)
            currentRow.Cells(1).SetValue(startDate.AddDays(i))
            currentRow.Cells(2).SetValue(random.Next(1, 12))
        Next

        ' Calculate formulas in worksheet.
        worksheet.Calculate()

        workbook.Save("Template Use.xlsx")
    End Sub
End Module