Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load Excel file with preservation feature enabled.
        Dim loadOptions As New XlsxLoadOptions() With {.PreserveUnsupportedFeatures = True}
        Dim workbook = ExcelFile.Load("Macros.xlsm", loadOptions)

        ' Save Excel file to output file of same format together with
        ' preserved information (unsupported features) from input file.
        workbook.Save("Preserved Output.xlsm")

    End Sub
End Module