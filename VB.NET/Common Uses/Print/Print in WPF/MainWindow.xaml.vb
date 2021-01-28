Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Xps.Packaging
Imports Microsoft.Win32
Imports GemBox.Spreadsheet

Partial Public Class MainWindow
    Inherits Window

    Dim workbook As ExcelFile

    Public Sub New()
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        InitializeComponent()
    End Sub

    Private Sub LoadFileBtn_Click(sender As Object, e As RoutedEventArgs)

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter =
            "XLSX files (*.xlsx, *.xltx, *.xlsm, *.xltm)|*.xlsx;*.xltx;*.xlsm;*.xltm" &
            "|XLS files (*.xls, *.xlt)|*.xls;*.xlt" &
            "|ODS files (*.ods, *.ots)|*.ods;*.ots" &
            "|CSV files (*.csv, *.tsv)|*.csv;*.tsv" &
            "|HTML files (*.html, *.htm)|*.html;*.htm"

        If (openFileDialog.ShowDialog() = True) Then
            Me.workbook = ExcelFile.Load(openFileDialog.FileName)
            Me.ShowPrintPreview()
        End If

    End Sub

    Private Sub PrintFileBtn_Click(sender As Object, e As RoutedEventArgs)

        If Me.workbook Is Nothing Then Return

        Dim printDialog As New PrintDialog() With {.UserPageRangeEnabled = True}
        If (printDialog.ShowDialog() = True) Then

            Dim printOptions As New PrintOptions(printDialog.PrintTicket.GetXmlStream()) With
            {
                .SelectionType = SelectionType.EntireFile
            }

            printOptions.FromPage = printDialog.PageRange.PageFrom - 1
            printOptions.ToPage = If(printDialog.PageRange.PageTo = 0, Integer.MaxValue, printDialog.PageRange.PageTo - 1)

            Me.workbook.Print(printDialog.PrintQueue.FullName, printOptions)
        End If

    End Sub

    Private Sub ShowPrintPreview()

        Dim xpsDocument As XpsDocument = workbook.ConvertToXpsDocument(
            New XpsSaveOptions() With {.SelectionType = SelectionType.EntireFile})

        ' Note, XpsDocument must stay referenced so that DocumentViewer can access additional resources from it.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
        Me.DocViewer.Tag = xpsDocument
        Me.DocViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

End Class