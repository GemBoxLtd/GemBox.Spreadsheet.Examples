Imports System
Imports System.Data
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile()
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("DataTable to Sheet")

        Dim dt As DataTable = New DataTable()

        dt.Columns.Add("ID", Type.GetType("System.Int32"))
        dt.Columns.Add("FirstName", Type.GetType("System.String"))
        dt.Columns.Add("LastName", Type.GetType("System.String"))

        dt.Rows.Add(New Object() {100, "John", "Doe"})
        dt.Rows.Add(New Object() {101, "Fred", "Nurk"})
        dt.Rows.Add(New Object() {103, "Hans", "Meier"})
        dt.Rows.Add(New Object() {104, "Ivan", "Horvat"})
        dt.Rows.Add(New Object() {105, "Jean", "Dupont"})
        dt.Rows.Add(New Object() {106, "Mario", "Rossi"})

        ws.Cells(0, 0).Value = "DataTable insert example:"

        ' Insert DataTable into an Excel worksheet.
        ws.InsertDataTable(dt,
            New InsertDataTableOptions() With
            {
                .ColumnHeaders = True,
                .StartRow = 2
            })

        ef.Save("DataTable to Sheet.xlsx")

    End Sub

End Module