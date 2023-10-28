Imports System.Data
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("DataTable to Sheet")

        Dim dataTable = New DataTable

        dataTable.Columns.Add("ID", Type.GetType("System.Int32"))
        dataTable.Columns.Add("FirstName", Type.GetType("System.String"))
        dataTable.Columns.Add("LastName", Type.GetType("System.String"))

        dataTable.Rows.Add(New Object() {100, "John", "Doe"})
        dataTable.Rows.Add(New Object() {101, "Fred", "Nurk"})
        dataTable.Rows.Add(New Object() {103, "Hans", "Meier"})
        dataTable.Rows.Add(New Object() {104, "Ivan", "Horvat"})
        dataTable.Rows.Add(New Object() {105, "Jean", "Dupont"})
        dataTable.Rows.Add(New Object() {106, "Mario", "Rossi"})

        worksheet.Cells(0, 0).Value = "DataTable insert example:"

        ' Insert DataTable to an Excel worksheet.
        worksheet.InsertDataTable(dataTable,
            New InsertDataTableOptions() With
            {
                .ColumnHeaders = True,
                .StartRow = 2
            })

        workbook.Save("DataTable to Sheet.xlsx")
    End Sub
End Module