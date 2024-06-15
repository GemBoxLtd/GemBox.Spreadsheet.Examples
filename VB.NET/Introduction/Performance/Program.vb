Imports BenchmarkDotNet.Attributes
Imports BenchmarkDotNet.Engines
Imports BenchmarkDotNet.Jobs
Imports BenchmarkDotNet.Running
Imports GemBox.Spreadsheet
Imports System.Collections.Generic
Imports System.IO

<SimpleJob(RuntimeMoniker.Net80)>
<SimpleJob(RuntimeMoniker.Net48)>
Public Class Program

    Private workbook As ExcelFile
    Private ReadOnly consumer As Consumer = New Consumer()

    Public Shared Sub Main()
        BenchmarkRunner.Run(Of Program)()
    End Sub

    <GlobalSetup>
    Public Sub SetLicense()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        ' https://www.gemboxsoftware.com/spreadsheet/examples/free-trial-professional/1001

        Me.workbook = ExcelFile.Load("RandomSheets.xlsx")
    End Sub

    <Benchmark>
    Public Function Reading() As ExcelFile
        Return ExcelFile.Load("RandomSheets.xlsx")
    End Function

    <Benchmark>
    Public Sub Writing()
        Using stream = New MemoryStream()
            Me.workbook.Save(stream, New XlsxSaveOptions())
        End Using
    End Sub

    <Benchmark>
    Public Sub Iterating()
        Me.LoopThroughAllCells().Consume(Me.consumer)
    End Sub

    Public Iterator Function LoopThroughAllCells() As IEnumerable(Of Object)
        For Each worksheet In Me.workbook.Worksheets
            For Each row In worksheet.Rows
                For Each cell In row.AllocatedCells
                    Yield cell.Value
                Next
            Next
        Next
    End Function

End Class
