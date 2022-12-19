Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Excel workbook from file's path.
        Dim workbook As ExcelFile = ExcelFile.Load("CombinedTemplate.xlsx")

        ' Set sheets print options.
        For Each worksheet As ExcelWorksheet In workbook.Worksheets
            Dim sheetPrintOptions As ExcelPrintOptions = worksheet.PrintOptions

            sheetPrintOptions.Portrait = False
            sheetPrintOptions.HorizontalCentered = True
            sheetPrintOptions.VerticalCentered = True

            sheetPrintOptions.PrintHeadings = True
            sheetPrintOptions.PrintGridlines = True
        Next

        ' Create spreadsheet's print options. 
        Dim printOptions As New PrintOptions()
        printOptions.SelectionType = SelectionType.EntireFile

        ' Print Excel workbook to default printer (e.g. 'Microsoft Print to Pdf').
        Dim printerName As String = Nothing
        workbook.Print(printerName, printOptions)

    End Sub
End Module