Imports GemBox.Spreadsheet
Imports System
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms

Partial Public Class Form1
    Inherits Form

    Dim workbook As ExcelFile

    Public Sub New()
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        InitializeComponent()
    End Sub

    Private Sub LoadFileMenuItem_Click(sender As Object, e As EventArgs) Handles LoadFileMenuItem.Click

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter =
            "XLSX files (*.xlsx, *.xltx, *.xlsm, *.xltm)|*.xlsx;*.xltx;*.xlsm;*.xltm" &
            "|XLS files (*.xls, *.xlt)|*.xls;*.xlt" &
            "|ODS files (*.ods, *.ots)|*.ods;*.ots" &
            "|CSV files (*.csv, *.tsv)|*.csv;*.tsv" &
            "|HTML files (*.html, *.htm)|*.html;*.htm"

        If (openFileDialog.ShowDialog() = DialogResult.OK) Then
            Me.workbook = ExcelFile.Load(openFileDialog.FileName)
            Me.ShowPrintPreview()
        End If

    End Sub

    Private Sub PrintFileMenuItem_Click(sender As Object, e As EventArgs) Handles PrintFileMenuItem.Click

        If Me.workbook Is Nothing Then Return

        Dim printDialog As New PrintDialog() With {.AllowSomePages = True}
        If (printDialog.ShowDialog() = DialogResult.OK) Then

            Dim printerSettings As PrinterSettings = printDialog.PrinterSettings
            Dim printOptions As New PrintOptions() With {.SelectionType = SelectionType.EntireFile}

            ' Set PrintOptions properties based on PrinterSettings properties.
            printOptions.CopyCount = printerSettings.Copies
            printOptions.FromPage = If(printerSettings.FromPage = 0, 0, printerSettings.FromPage - 1)
            printOptions.ToPage = If(printerSettings.ToPage = 0, Integer.MaxValue, printerSettings.ToPage - 1)

            Me.workbook.Print(printerSettings.PrinterName, printOptions)
        End If

    End Sub

    Private Sub ShowPrintPreview()

        ' Create image for each Excel workbook's page.
        Dim images As Image() = Me.CreatePrintPreviewImages()
        Dim imageIndex As Integer = 0

        ' Draw each page's image on PrintDocument for print preview.
        Dim printDocument = New PrintDocument()
        AddHandler printDocument.PrintPage,
            Sub(sender, e)
                Using image As Image = images(imageIndex)
                    Dim graphics = e.Graphics
                    Dim region = graphics.VisibleClipBounds

                    ' Rotate image if it has landscape orientation.
                    If image.Width > image.Height Then image.RotateFlip(RotateFlipType.Rotate270FlipNone)

                    graphics.DrawImage(image, 0, 0, region.Width, region.Height)
                End Using

                imageIndex += 1
                e.HasMorePages = imageIndex < images.Length
            End Sub

        Me.PageUpDown.Value = 1
        Me.PageUpDown.Maximum = images.Length
        Me.printPreviewControl.Document = printDocument

    End Sub

    Private Function CreatePrintPreviewImages() As Image()

        Dim paginatorOptions As New PaginatorOptions With {.SelectionType = SelectionType.EntireFile}
        Dim pages = Me.workbook.GetPaginator(paginatorOptions).Pages

        Dim images = New Image(pages.Count - 1) {}
        Dim imageOptions As New ImageSaveOptions()

        For pageIndex As Integer = 0 To pages.Count - 1
            Dim imageStream = New MemoryStream()
            pages(pageIndex).Save(imageStream, imageOptions)
            images(pageIndex) = Image.FromStream(imageStream)
        Next

        Return images

    End Function

    Private Sub PageUpDown_ValueChanged(sender As Object, e As EventArgs) Handles PageUpDown.ValueChanged
        Me.printPreviewControl.StartPage = Me.PageUpDown.Value - 1
    End Sub

End Class