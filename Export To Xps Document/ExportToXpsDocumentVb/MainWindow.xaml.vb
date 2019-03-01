Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Xps.Packaging
Imports GemBox.Spreadsheet

Class MainWindow

    Dim xpsDocument As XpsDocument

    Public Sub New()

        InitializeComponent()

        SetDocumentViewer(Me.DocumentViewer)
    End Sub

    Private Sub SetDocumentViewer(documentViewer As DocumentViewer)

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile

        Dim worksheet = workbook.Worksheets.Add("Sheet1")

        worksheet.Cells(0, 0).Value = "English:"
        worksheet.Cells(0, 1).Value = "Hello"

        worksheet.Cells(1, 0).Value = "Russian:"
        worksheet.Cells(1, 1).Value = New String(New Char() {ChrW(&H417), ChrW(&H434), ChrW(&H440), ChrW(&H430), ChrW(&H432), ChrW(&H441), ChrW(&H442), ChrW(&H432), ChrW(&H443), ChrW(&H439), ChrW(&H442), ChrW(&H435)})

        worksheet.Cells(2, 0).Value = "Chinese:"
        worksheet.Cells(2, 1).Value = New String(New Char() {ChrW(&H4F60), ChrW(&H597D)})

        worksheet.Cells(4, 0).Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC."
        worksheet.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = True

        worksheet.HeadersFooters.DefaultPage.Header.CenterSection.Content = "Export To XpsDocument / DocumentViewer Control Example"

        worksheet.PrintOptions.PrintGridlines = True

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        xpsDocument = workbook.ConvertToXpsDocument(SaveOptions.XpsDefault)

        documentViewer.Document = xpsDocument.GetFixedDocumentSequence()
    End Sub
End Class