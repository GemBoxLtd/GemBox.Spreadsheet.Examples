Imports System.IO
Imports GemBox.Spreadsheet

Public NotInheritable Class LargeCsvReader
    Inherits TextReader

    Private Const MaxRow As Integer = 1_048_576
    Private ReadOnly reader As TextReader
    Private ReadOnly options As CsvLoadOptions

    Private currentRow As Integer
    Private finished As Boolean

    Public Shared Function ReadFile(path As String, options As CsvLoadOptions) As ExcelFile
        Dim workbook As New ExcelFile()
        Dim sheetIndex As Integer = 0

        Using reader = New LargeCsvReader(path, options)
            While reader.CanReadNextSheet()
                sheetIndex += 1
                reader.ReadSheet(workbook, $"Sheet{sheetIndex}")
            End While
        End Using

        Return workbook
    End Function

    Private Sub New(path As String, options As CsvLoadOptions)
        Me.reader = File.OpenText(path)
        Me.options = options
    End Sub

    Public Overrides Function ReadLine() As String
        If Me.currentRow = MaxRow Then Return Nothing

        Me.currentRow += 1
        Dim line As String = Me.reader.ReadLine()
        If line Is Nothing Then Me.finished = True

        Return line
    End Function

    Private Sub ReadSheet(ByVal workbook As ExcelFile, ByVal name As String)
        Dim worksheet = ExcelFile.Load(Me, Me.options).Worksheets.ActiveWorksheet
        workbook.Worksheets.AddCopy(name, worksheet)
    End Sub

    Private Function CanReadNextSheet() As Boolean
        If Me.finished Then Return False

        Me.currentRow = 0
        Return True
    End Function

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Me.reader.Dispose()
    End Sub
End Class