Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Print and View Options")

        ws.Cells("M1").Value = "This worksheet shows how to set various print related and view related options."
        ws.Cells.GetSubrange("G1", "M1").Merged = True
        ws.Cells("M2").Value = "To see results of print options, go to Print and Page Setup dialogs in MS Excel."
        ws.Cells.GetSubrange("G2", "M2").Merged = True
        ws.Cells("M3").Value = "Notice that print and view options are worksheet based, not workbook based."
        ws.Cells.GetSubrange("G3", "M3").Merged = True

        ' Print options:
        Dim printOptions = ws.PrintOptions
        printOptions.PrintGridlines = True
        printOptions.PrintHeadings = True
        printOptions.Portrait = False
        printOptions.PaperType = PaperType.A3
        printOptions.NumberOfCopies = 5

        ' View options:
        ws.ViewOptions.FirstVisibleColumn = 3
        ws.ViewOptions.ShowColumnsFromRightToLeft = True
        ws.ViewOptions.Zoom = 123

        ' Set print area
        ws.NamedRanges.SetPrintArea(ws.Cells.GetSubrange("E1", "U7"))

        ef.Save("Print and View Options.xlsx")

    End Sub

End Module