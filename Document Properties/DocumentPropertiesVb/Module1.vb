Imports System.Collections.Generic
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("TemplateUse.xlsx")

        ' Add Sheet
        Dim ws As ExcelWorksheet = ef.Worksheets.InsertEmpty(0, "Document Properties")
        ef.Worksheets.ActiveWorksheet = ws

        Dim rowIndex As Integer = 0
        ' Read Built-in Document Properies 
        ws.Cells(rowIndex, 0).Value = "Built-in document properties"
        rowIndex = rowIndex + 1

        ws.Cells(rowIndex, 0).Value = "Property"
        ws.Cells(rowIndex, 1).Value = "Value"
        rowIndex = rowIndex + 1

        For Each keyValue As KeyValuePair(Of BuiltInDocumentProperties, String) In ef.DocumentProperties.BuiltIn

            ws.Cells(rowIndex, 0).Value = keyValue.Key.ToString()
            ws.Cells(rowIndex, 1).Value = keyValue.Value
            rowIndex = rowIndex + 1

        Next

        ' Read Custom Document Properties
        rowIndex = rowIndex + 1
        ws.Cells(rowIndex, 0).Value = "Custom Document Properties"

        rowIndex = rowIndex + 1
        ws.Cells(rowIndex, 0).Value = "Property"
        ws.Cells(rowIndex, 1).Value = "Value"
        rowIndex = rowIndex + 1

        For Each keyValue As KeyValuePair(Of String, Object) In ef.DocumentProperties.Custom

            ws.Cells(rowIndex, 0).Value = keyValue.Key
            ws.Cells(rowIndex, 1).Value = keyValue.Value.ToString()
            rowIndex = rowIndex + 1

        Next

        ' Write/Modifiy Document Properties
        ef.DocumentProperties.BuiltIn(BuiltInDocumentProperties.Author) = "John Doe"
        ef.DocumentProperties.BuiltIn(BuiltInDocumentProperties.Title) = "Genrated title"

        ws.Columns(0).AutoFit()
        ws.Columns(1).AutoFit()

        ef.Save("Document Properties.xlsx")

    End Sub

End Module