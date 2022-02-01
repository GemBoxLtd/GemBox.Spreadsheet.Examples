Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Template.xlsx")

        ' Get template sheet.
        Dim templateSheet = workbook.Worksheets(0)

        ' Copy template sheet.
        For i = 0 To 3
            workbook.Worksheets.AddCopy("Invoice " + (i + 1).ToString(), templateSheet)
        Next

        ' Delete template sheet.
        workbook.Worksheets.Remove(0)

        Dim random As New Random()

        ' For each sheet.
        For i = 0 To 3

            ' Get sheet.
            Dim worksheet = workbook.Worksheets(i)

            ' Write sheet's cells.
            worksheet.Cells("C6").Value = "ACME Corp"
            worksheet.Cells("C7").Value = "240 Old Country Road, Springfield, IL"

            Dim startDate As DateTime = DateTime.Today
            Dim itemsCount As Integer = random.Next(5, 20)
            worksheet.Cells("C11").SetValue(startDate)
            worksheet.Cells("C12").SetValue(startDate.AddDays(itemsCount - 1))

            ' Copy template row.
            Dim row As Integer = 17
            worksheet.Rows.InsertCopy(row + 1, itemsCount - 1, worksheet.Rows(row))

            ' Write row's cells.
            For j = 0 To itemsCount - 1
                Dim currentRow = worksheet.Rows(row + j)
                currentRow.Cells(1).SetValue(startDate.AddDays(j))
                currentRow.Cells(2).SetValue(random.Next(6, 9))
            Next

        Next

        workbook.Save("Sheet Copying_Deleting.xlsx")

    End Sub
End Module