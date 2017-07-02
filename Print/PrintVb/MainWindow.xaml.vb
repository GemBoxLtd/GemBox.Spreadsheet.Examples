Imports Microsoft.Win32
Imports GemBox.Spreadsheet

Class MainWindow

    Dim ef As ExcelFile

    Public Sub New()
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        InitializeComponent()

        Me.EnableControls()
    End Sub

    Private Sub LoadFileBtn_Click(sender As Object, e As RoutedEventArgs)

        Dim fileDialog = New OpenFileDialog()
        fileDialog.Filter = "XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|XLS files (*.xls, *.xlt)|*.xls;*.xlt|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm"

        If (fileDialog.ShowDialog() = True) Then
            Me.ef = ExcelFile.Load(fileDialog.FileName)

            Me.ShowPrintPreview()
            Me.EnableControls()
        End If

    End Sub

    Private Sub SimplePrint_Click(sender As Object, e As RoutedEventArgs)

        ' Print to default printer using default options
        Me.ef.Print()

    End Sub

    Private Sub AdvancedPrint_Click(sender As Object, e As RoutedEventArgs)

        ' We can use PrintDialog for defining print options
        Dim printDialog = New PrintDialog()
        printDialog.UserPageRangeEnabled = True

        If (printDialog.ShowDialog() = True) Then

            Dim printOptions = New PrintOptions(printDialog.PrintTicket.GetXmlStream())

            printOptions.FromPage = printDialog.PageRange.PageFrom - 1
            If (printDialog.PageRange.PageTo = 0) Then
                printOptions.ToPage = Int32.MaxValue
            Else
                printOptions.ToPage = printDialog.PageRange.PageTo - 1
            End If

            Me.ef.Print(printDialog.PrintQueue.FullName, printOptions)
        End If

    End Sub

    ' We can use DocumentViewer for print preview (but we don't need).
    Private Sub ShowPrintPreview()

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        Dim xpsDocument = ef.ConvertToXpsDocument(SaveOptions.XpsDefault)
        Me.DocViewer.Tag = xpsDocument

        Me.DocViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

    Private Sub EnableControls()

        Dim isEnabled As Boolean = Me.ef IsNot Nothing

        Me.DocViewer.IsEnabled = isEnabled
        Me.SimplePrintFileBtn.IsEnabled = isEnabled
        Me.AdvancedPrintFileBtn.IsEnabled = isEnabled

    End Sub

End Class
