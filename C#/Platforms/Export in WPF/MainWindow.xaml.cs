using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Xps.Packaging;
using GemBox.Spreadsheet;
using Microsoft.Win32;

partial class MainWindow : Window
{
    private ExcelFile workbook;
    private XpsDocument xpsDocument;
    private ImageSource imageSource;

    private Action updateSourceAction;

    public MainWindow()
    {
        this.InitializeComponent();

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        this.updateSourceAction = this.SetImageSource;

        this.InitExcelFile();
        this.updateSourceAction();
    }

    private void InitExcelFile()
    {
        this.workbook = new ExcelFile();
        var worksheet = this.workbook.Worksheets.Add("Sheet1");

        worksheet.Cells[0, 0].Value = "English:";
        worksheet.Cells[0, 1].Value = "Hello";

        worksheet.Cells[1, 0].Value = "Russian:";
        worksheet.Cells[1, 1].Value = new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' });

        worksheet.Cells[2, 0].Value = "Chinese:";
        worksheet.Cells[2, 1].Value = new string(new char[] { '\u4f60', '\u597d' });

        worksheet.Cells[4, 0].Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC.";
        worksheet.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = true;

        worksheet.HeadersFooters.DefaultPage.Header.CenterSection.Content = "Export To ImageSource / Image Control Example";

        worksheet.PrintOptions.PrintGridlines = true;
    }

    private void SetImageSource()
    {
        if (this.imageSource == null)
            this.imageSource = this.workbook.ConvertToImageSource(SaveOptions.ImageDefault);

        this.DocumentViewer.Document = null;
        this.ImageControl.Source = this.imageSource;

        this.DocumentViewer.Visibility = Visibility.Collapsed;
        this.ImageScrollViewer.Visibility = Visibility.Visible;
    }

    private void SetDocumentViewerSource()
    {
        // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        if (this.xpsDocument == null)
            this.xpsDocument = this.workbook.ConvertToXpsDocument(SaveOptions.XpsDefault);

        this.ImageControl.Source = null;
        this.DocumentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();

        this.ImageScrollViewer.Visibility = Visibility.Collapsed;
        this.DocumentViewer.Visibility = Visibility.Visible;
    }

    private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
        };
        if (openFileDialog.ShowDialog() != true)
            return;

        this.workbook = ExcelFile.Load(openFileDialog.FileName);
        this.xpsDocument = null;
        this.imageSource = null;
        this.updateSourceAction();
    }

    private void BtnShowAsImage_Click(object sender, RoutedEventArgs e)
    {
        this.updateSourceAction = this.SetImageSource;
        this.updateSourceAction();
    }

    private void BtnShowAsDocument_Click(object sender, RoutedEventArgs e)
    {
        this.updateSourceAction = this.SetDocumentViewerSource;
        this.updateSourceAction();
    }
}