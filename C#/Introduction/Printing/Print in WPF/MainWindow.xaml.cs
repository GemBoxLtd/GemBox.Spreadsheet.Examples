using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using Microsoft.Win32;
using GemBox.Spreadsheet;

public partial class MainWindow : Window
{
    private ExcelFile workbook;

    public MainWindow()
    {
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private void LoadFileBtn_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter =
            "XLSX files (*.xlsx, *.xltx, *.xlsm, *.xltm)|*.xlsx;*.xltx;*.xlsm;*.xltm" +
            "|XLS files (*.xls, *.xlt)|*.xls;*.xlt" +
            "|ODS files (*.ods, *.ots)|*.ods;*.ots" +
            "|CSV files (*.csv, *.tsv)|*.csv;*.tsv" +
            "|HTML files (*.html, *.htm)|*.html;*.htm";

        if (openFileDialog.ShowDialog() == true)
        {
            this.workbook = ExcelFile.Load(openFileDialog.FileName);
            this.ShowPrintPreview();
        }
    }

    private void PrintFileBtn_Click(object sender, RoutedEventArgs e)
    {
        if (this.workbook == null)
            return;

        PrintDialog printDialog = new PrintDialog() { UserPageRangeEnabled = true };
        if (printDialog.ShowDialog() == true)
        {
            PrintOptions printOptions = new PrintOptions(printDialog.PrintTicket.GetXmlStream())
            {
                SelectionType = SelectionType.EntireFile
            };

            printOptions.FromPage = printDialog.PageRange.PageFrom - 1;
            printOptions.ToPage = printDialog.PageRange.PageTo == 0 ? int.MaxValue : printDialog.PageRange.PageTo - 1;

            this.workbook.Print(printDialog.PrintQueue.FullName, printOptions);
        }
    }

    private void ShowPrintPreview()
    {
        XpsDocument xpsDocument = this.workbook.ConvertToXpsDocument(
            new XpsSaveOptions() { SelectionType = SelectionType.EntireFile });

        // Note, XpsDocument must stay referenced so that DocumentViewer can access additional resources from it.
        // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will no longer work.
        this.DocViewer.Tag = xpsDocument;
        this.DocViewer.Document = xpsDocument.GetFixedDocumentSequence();
    }
}