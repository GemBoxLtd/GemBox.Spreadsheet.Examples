Imports System
Imports System.Windows.Forms
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.WinFormsUtilities

Public Class Form1

    Public Sub New()

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        InitializeComponent()
    End Sub

    Private Sub btnLoadFile_Click(sender As Object, e As EventArgs) Handles btnLoadFile.Click

        Dim openFileDialog = New OpenFileDialog()
        openFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm"
        openFileDialog.FilterIndex = 2

        If (openFileDialog.ShowDialog() = DialogResult.OK) Then

            Dim workbook = ExcelFile.Load(openFileDialog.FileName)

            ' From ExcelFile to DataGridView.
            DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, Me.dataGridView1, New ExportToDataGridViewOptions() With {.ColumnHeaders = True})
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim saveFileDialog = New SaveFileDialog()
        saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp"
        saveFileDialog.FilterIndex = 3

        If (saveFileDialog.ShowDialog() = DialogResult.OK) Then

            Dim workbook = New ExcelFile()
            Dim worksheet = workbook.Worksheets.Add("Sheet1")

            ' From DataGridView to ExcelFile.
            DataGridViewConverter.ImportFromDataGridView(worksheet, Me.dataGridView1, New ImportFromDataGridViewOptions() With {.ColumnHeaders = True})

            workbook.Save(saveFileDialog.FileName)
        End If
    End Sub
End Class