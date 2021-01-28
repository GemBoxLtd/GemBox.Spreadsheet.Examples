Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' In order to convert Excel to PDF, we just need to
        '   1. Load XLS Or XLSX file into ExcelFile object.
        '   2. Save ExcelFile object to PDF file.
        Dim workbook As ExcelFile = ExcelFile.Load("ComplexTemplate.xlsx")
        workbook.Save("Convert1.pdf", New PdfSaveOptions() With {.SelectionType = SelectionType.EntireFile})
    End Sub

    Sub Example2()
        ' Load Excel file.
        Dim workbook As ExcelFile = ExcelFile.Load("ComplexTemplate.xlsx")

        ' Get Excel sheet you want to export.
        Dim worksheet As ExcelWorksheet = workbook.Worksheets(0)

        ' Set targeted sheet as active.
        workbook.Worksheets.ActiveWorksheet = worksheet

        ' Get cell range that you want to export.
        Dim range As CellRange = worksheet.Cells.GetSubrange("A5:I14")

        ' Set targeted range as print area.
        worksheet.NamedRanges.SetPrintArea(range)

        ' Save to PDF file.
        ' By default, the SelectionType.ActiveSheet is used.
        workbook.Save("Convert2.pdf")
    End Sub

    Sub Example3()
        Dim conformanceLevel As PdfConformanceLevel = PdfConformanceLevel.PdfA1a

        ' Load Excel file.
        Dim workbook = ExcelFile.Load("ComplexTemplate.xlsx")

        ' Create PDF save options.
        Dim options As New PdfSaveOptions() With
        {
            .ConformanceLevel = conformanceLevel
        }

        ' Save to PDF file.
        workbook.Save("Output3.pdf", options)
    End Sub

End Module