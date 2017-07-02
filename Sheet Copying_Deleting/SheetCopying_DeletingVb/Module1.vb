Imports System
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("TemplateUse.xlsx")

        ' Get template sheet.
        Dim templateSheet As ExcelWorksheet = ef.Worksheets(0)

        ' Copy template sheet.
        Dim i As Int32
        For i = 0 To 3 Step 1
            ef.Worksheets.AddCopy("Invoice " + (i + 1).ToString(), templateSheet)
        Next

        ' Delete template sheet.
        ef.Worksheets.Remove(0)

        Dim startTime As DateTime = DateTime.Now

        ' Go to the first Monday from today.
        While startTime.DayOfWeek <> DayOfWeek.Monday
            startTime = startTime.AddDays(1)
        End While

        Dim rnd As Random = New Random()

        ' For each sheet.
        For i = 0 To 3 Step 1

            ' Get sheet.
            Dim ws As ExcelWorksheet = ef.Worksheets(i)

            ' Set some fields.
            ws.Cells("J5").SetValue(14 + i)
            ws.Cells("J6").SetValue(DateTime.Now)
            ws.Cells("J6").Style.NumberFormat = "m/dd/yyyy"

            ws.Cells("D12").Value = "ACME Corp"
            ws.Cells("D13").Value = "240 Old Country Road, Springfield, IL"
            ws.Cells("D14").Value = "USA"
            ws.Cells("D15").Value = "Joe Smith"

            ws.Cells("E18").Value = String.Format(startTime.ToShortDateString() + " until " + startTime.AddDays(11).ToShortDateString())

            Dim j As Int32
            For j = 0 To 9 Step 1
                ws.Cells(21 + j, 1).SetValue(startTime) ' Set date.
                ws.Cells(21 + j, 1).Style.NumberFormat = "dddd, mmmm dd, yyyy"
                ws.Cells(21 + j, 4).SetValue(rnd.Next(6, 9)) ' Work hours.

                ' Skip Saturday and Sunday.
                If j = 4 Then startTime = startTime.AddDays(3) Else startTime = startTime.AddDays(1)
            Next

            ' Skip Saturday and Sunday.
            startTime = startTime.AddDays(2)

            ws.Cells("B36").Value = "Payment via check."

        Next

        ef.Save("Sheet Copying_Deleting.xlsx")

    End Sub

End Module