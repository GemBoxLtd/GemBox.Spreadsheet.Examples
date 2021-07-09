Imports System
Imports System.IO
Imports System.Text
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Types")

        worksheet.Rows(0).Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Columns(0).Width = 25 * 256
        worksheet.Columns(1).Width = 25 * 256
        worksheet.Columns(2).Width = 25 * 256

        worksheet.Cells(0, 0).Value = "Value"
        worksheet.Cells(0, 1).Value = ".NET Value Type"
        worksheet.Cells(0, 2).Value = "Cell Value Type"

        ' Sample data values.
        Dim values As Object() =
        {
            DBNull.Value,
            Byte.MaxValue,
            SByte.MinValue,
            Short.MinValue,
            UShort.MaxValue,
            1000,
            CUInt(2000),
            Long.MinValue,
            ULong.MaxValue,
            Single.MaxValue,
            Double.MaxValue,
            3000.45D,
            True,
            DateTime.Now,
            "a"c,
            "Sample text.",
            New StringBuilder("Sample text.")
        }

        ' Write data and data type to Excel cells.
        For i = 0 To values.Length - 1
            Dim value As Object = values(i)
            worksheet.Cells(i + 1, 0).Value = value
            worksheet.Cells(i + 1, 1).Value = value.GetType().ToString()
        Next

        ' Save to Excel file and load it back as ExcelFile object.
        Using stream As New MemoryStream()
            workbook.Save(stream, SaveOptions.XlsxDefault)
            workbook = ExcelFile.Load(stream, LoadOptions.XlsxDefault)
            worksheet = workbook.Worksheets(0)
        End Using

        ' Write cell type to Excel cells.
        For i = 0 To values.Length - 1
            worksheet.Cells(i + 1, 2).Value = worksheet.Cells(i + 1, 0).ValueType.ToString()
        Next

        workbook.Save("Data Types.xlsx")

    End Sub
End Module