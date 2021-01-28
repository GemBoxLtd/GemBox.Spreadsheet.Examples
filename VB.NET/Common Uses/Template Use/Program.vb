Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim numberOfItems As Integer = 10
        Dim startDate = DateTime.Today.AddDays(-numberOfItems)
        Dim endDate = DateTime.Today

        ' Load an Excel template.
        Dim workbook = ExcelFile.Load("Template.xlsx")

        ' Get template sheet.
        Dim worksheet = workbook.Worksheets(0)

        ' Find cells with placeholder text and set their values.
        Dim row As Integer, column As Integer
        If worksheet.Cells.FindText("[Company Name]", row, column) Then
            worksheet.Cells(row, column).Value = "ACME Corp"
        End If
        If worksheet.Cells.FindText("[Company Address]", row, column) Then
            worksheet.Cells(row, column).Value = "240 Old Country Road, Springfield, IL"
        End If
        If worksheet.Cells.FindText("[Start Date]", row, column) Then
            worksheet.Cells(row, column).Value = startDate
        End If
        If worksheet.Cells.FindText("[End Date]", row, column) Then
            worksheet.Cells(row, column).Value = endDate
        End If

        ' Copy template row.
        row = 17
        worksheet.Rows.InsertCopy(row + 1, numberOfItems - 1, worksheet.Rows(row))

        ' Fill copied rows with sample data.
        Dim random As New Random()
        For i As Integer = 0 To numberOfItems - 1
            Dim currentRow = worksheet.Rows(row + i)
            currentRow.Cells(1).SetValue(startDate.AddDays(i))
            currentRow.Cells(2).SetValue(random.Next(1, 12))
        Next

        ' Calculate formulas in a sheet.
        worksheet.Calculate()

        ' Save the modified Excel template to output file.
        workbook.Save("Output.xlsx")

    End Sub
End Module