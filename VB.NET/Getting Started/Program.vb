Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Hello World")

        worksheet.Cells(0, 0).Value = "English:"
        worksheet.Cells(0, 1).Value = "Hello"

        ' Use UNICODE string.
        worksheet.Cells(1, 0).Value = "Ukrainian:"
        worksheet.Cells(1, 1).Value = "Привіт"

        ' Use UNICODE characters.
        worksheet.Cells(2, 0).Value = "Chinese:"
        worksheet.Cells(2, 1).Value = New String(New Char() {ChrW(&H4F60), ChrW(&H597D)})

        worksheet.Cells(4, 0).Value =
            "In order to see Ukrainian and Chinese characters " &
            "you need to have appropriate fonts on your PC."

        worksheet.Cells.GetSubrange("A5", "K5").Merged = True

        workbook.Save("HelloWorld.xlsx")

    End Sub

End Module