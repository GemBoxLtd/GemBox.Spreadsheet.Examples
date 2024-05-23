Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Tables

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Tables")

        ' Add some data.
        Dim data(,) = New Object(4, 2) _
        {
            {"Worker", "Hours", "Price"},
            {"John Doe", 25, 35.0},
            {"Jane Doe", 27, 35.0},
            {"Jack White", 18, 32.0},
            {"George Black", 31, 35.0}
        }

        For i As Integer = 0 To 4
            For j As Integer = 0 To 2
                worksheet.Cells.Item(i, j).Value = data(i, j)
            Next
        Next

        ' Set column widths and formats.
        worksheet.Columns(0).SetWidth(100, LengthUnit.Pixel)
        worksheet.Columns(1).SetWidth(70, LengthUnit.Pixel)
        worksheet.Columns(2).SetWidth(70, LengthUnit.Pixel)
        worksheet.Columns(3).SetWidth(70, LengthUnit.Pixel)
        worksheet.Columns(2).Style.NumberFormat = """$""#,##0.00"
        worksheet.Columns(3).Style.NumberFormat = """$""#,##0.00"

        ' Create table And enable totals row.
        Dim table = worksheet.Tables.Add("Table1", "A1:C5", True)
        table.HasTotalsRow = True

        ' Add New column.
        Dim column = table.Columns.Add()
        column.Name = "Total"

        ' Populate column.
        For Each cell In column.DataRange
            cell.Formula = "=Table1[Hours] * Table1[Price]"
        Next

        ' Set totals row function for newly added column and calculate it.
        column.TotalsRowFunction = TotalsRowFunction.Sum
        column.Range.Calculate()

        ' Set table style.
        table.BuiltInStyle = BuiltInTableStyleName.TableStyleMedium2

        workbook.Save("Tables.xlsx")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Tables.xlsx")
        Dim worksheet = workbook.Worksheets("Tables")
        Dim table = worksheet.Tables("Table1")

        ' Remove existing table row.
        table.Rows.RemoveAt(0)

        ' Update existing table row.
        Dim tableRow = table.Rows(0)
        tableRow.DataRange(0).Value = "Jane Updated"
        tableRow.DataRange(1).Value = 30
        tableRow.DataRange(2).Value = 40.0

        ' Sample data for writing into a table.
        Dim data = {
            New Object() {"Fred Nurk", 22, 35.0},
            New Object() {"Hans Meier", 16, 20.0},
            New Object() {"Ivan Horvat", 24, 34.0}
        }

        For Each items As Object() In data
            ' Add new table row by adding cell values directly.
            tableRow = table.Rows.Add(items)
            tableRow.DataRange(3).Formula = "=Table1[Hours] * Table1[Price]"
        Next

        table.Columns("Total").Range.Calculate()

        workbook.Save("Tables Updated.xlsx")
    End Sub

End Module
