Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        ' Use Trial Mode
        AddHandler SpreadsheetInfo.FreeLimitReached,
            Sub(eventSender, args)
                args.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial
            End Sub

        Console.WriteLine("Creating file")

        ' Create large workbook
        Dim workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("sheet")
        For i As Integer = 0 To 1000000
            worksheet.Cells(i, 0).Value = i
        Next

        ' Create save options
        Dim saveOptions = New XlsxSaveOptions()
        AddHandler saveOptions.ProgressChanged,
            Sub(eventSender, args)
                Console.WriteLine($"Progress changed - {args.ProgressPercentage}%")
            End Sub

        ' Save file
        workbook.Save("file.xlsx", saveOptions)
    End Sub
End Module