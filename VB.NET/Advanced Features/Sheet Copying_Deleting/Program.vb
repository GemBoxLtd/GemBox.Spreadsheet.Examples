Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("TemplateUse.xlsx")

        ' Get template sheet.
        Dim templateSheet = workbook.Worksheets(0)

        ' Copy template sheet.
        For i = 0 To 3 Step 1
            workbook.Worksheets.AddCopy("Invoice " + (i + 1).ToString(), templateSheet)
        Next

        ' Delete template sheet.
        workbook.Worksheets.Remove(0)

        Dim startTime = DateTime.Now

        ' Go to the first Monday from today.
        While startTime.DayOfWeek <> DayOfWeek.Monday
            startTime = startTime.AddDays(1)
        End While

        Dim random As Random = New Random()

        ' For each sheet.
        For i = 0 To 3 Step 1

            ' Get sheet.
            Dim worksheet = workbook.Worksheets(i)

            ' Set some fields.
            worksheet.Cells("J5").SetValue(14 + i)
            worksheet.Cells("J6").SetValue(DateTime.Now)
            worksheet.Cells("J6").Style.NumberFormat = "m/dd/yyyy"

            worksheet.Cells("D12").Value = "ACME Corp"
            worksheet.Cells("D13").Value = "240 Old Country Road, Springfield, IL"
            worksheet.Cells("D14").Value = "USA"
            worksheet.Cells("D15").Value = "Joe Smith"

            worksheet.Cells("E18").Value = String.Format(startTime.ToShortDateString() + " until " + startTime.AddDays(11).ToShortDateString())

            For j = 0 To 9 Step 1

                worksheet.Cells(21 + j, 1).SetValue(startTime) ' Set date.
                worksheet.Cells(21 + j, 1).Style.NumberFormat = "dddd, mmmm dd, yyyy"
                worksheet.Cells(21 + j, 4).SetValue(random.Next(6, 9)) ' Work hours.

                ' Skip Saturday and Sunday.
                If j = 4 Then startTime = startTime.AddDays(3) Else startTime = startTime.AddDays(1)
            Next

            ' Skip Saturday and Sunday.
            startTime = startTime.AddDays(2)

            worksheet.Cells("B36").Value = "Payment via check."
        Next

        workbook.Save("Sheet Copying_Deleting.xlsx")
    End Sub
End Module