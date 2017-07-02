Imports System
Imports System.Text
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = ExcelFile.Load("SimpleTemplate.xlsx")

        Dim searchText = "Apollo 13"

        Dim ws = ef.Worksheets(0)

        Dim sb = New StringBuilder()

        Dim objectRow, objectColumn As Integer
        ws.Cells.FindText(searchText, False, False, objectRow, objectColumn)

        If objectRow = -1 Or objectColumn = -1 Then
            sb.AppendLine("Can't find text.")
        Else

            sb.AppendLine(searchText & " was launched on " & ws.Cells(objectRow, 2).Value & ".")

            Dim nationality = CType(ws.Cells(objectRow, 1).Value, String)

            If Not String.IsNullOrEmpty(nationality) Then

                Dim nationalityText = nationality.Trim().ToLowerInvariant()

                Dim nationalityCounter = 0

                Dim enumerator = ws.Columns(1).Cells.GetReadEnumerator()
                While enumerator.MoveNext()

                    Dim cell = enumerator.Current
                    Dim cellValue = CType(cell.Value, String)
                    If Not String.IsNullOrEmpty(cellValue) Then
                        If cellValue.Trim().ToLowerInvariant() = nationalityText Then nationalityCounter = nationalityCounter + 1
                    End If

                End While

                sb.AppendFormat("There are {0} entires for {1}.", nationalityCounter, nationality)
            End If

        End If

        Console.WriteLine(sb.ToString())

    End Sub

End Module