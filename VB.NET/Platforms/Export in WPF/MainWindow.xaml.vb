Imports GemBox.Spreadsheet
Imports Microsoft.Win32
Imports System
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Xps.Packaging

Class MainWindow
    Inherits Window
    Private workbook As ExcelFile
    Private xpsDocument As XpsDocument
    Private imageSource As ImageSource

    Private updateSourceAction As Action

    Public Sub New()
        Me.InitializeComponent()

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        updateSourceAction = AddressOf SetImageSource

        InitExcelFile()
        updateSourceAction()
    End Sub

    Private Sub InitExcelFile()
        workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Sheet1")

        worksheet.Cells(0, 0).Value = "English:"
        worksheet.Cells(0, 1).Value = "Hello"

        worksheet.Cells(1, 0).Value = "Russian:"
        worksheet.Cells(1, 1).Value = New String(New Char() {ChrW(&H417), ChrW(&H434), ChrW(&H440), ChrW(&H430), ChrW(&H432), ChrW(&H441), ChrW(&H442), ChrW(&H432), ChrW(&H443), ChrW(&H439), ChrW(&H442), ChrW(&H435)})

        worksheet.Cells(2, 0).Value = "Chinese:"
        worksheet.Cells(2, 1).Value = New String(New Char() {ChrW(&H4F60), ChrW(&H597D)})

        worksheet.Cells(4, 0).Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC."
        worksheet.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = True

        worksheet.HeadersFooters.DefaultPage.Header.CenterSection.Content = "Export To ImageSource / Image Control Example"

        worksheet.PrintOptions.PrintGridlines = True
    End Sub

    Private Sub SetImageSource()
        If imageSource Is Nothing Then imageSource = workbook.ConvertToImageSource(SaveOptions.ImageDefault)

        Me.DocumentViewer.Document = Nothing
        Me.ImageControl.Source = imageSource

        Me.DocumentViewer.Visibility = Visibility.Collapsed
        Me.ImageScrollViewer.Visibility = Visibility.Visible
    End Sub

    Private Sub SetDocumentViewerSource()
        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        If xpsDocument Is Nothing Then xpsDocument = workbook.ConvertToXpsDocument(SaveOptions.XpsDefault)

        Me.ImageControl.Source = Nothing
        Me.DocumentViewer.Document = xpsDocument.GetFixedDocumentSequence()

        Me.ImageScrollViewer.Visibility = Visibility.Collapsed
        Me.DocumentViewer.Visibility = Visibility.Visible
    End Sub

    Private Sub BtnOpenFile_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim openFileDialog = New OpenFileDialog With {
    .Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
}
        If openFileDialog.ShowDialog() <> True Then Return

        workbook = ExcelFile.Load(openFileDialog.FileName)
        xpsDocument = Nothing
        imageSource = Nothing
        updateSourceAction()
    End Sub

    Private Sub BtnShowAsImage_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        updateSourceAction = AddressOf SetImageSource
        updateSourceAction()
    End Sub

    Private Sub BtnShowAsDocument_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        updateSourceAction = AddressOf SetDocumentViewerSource
        updateSourceAction()
    End Sub
End Class
