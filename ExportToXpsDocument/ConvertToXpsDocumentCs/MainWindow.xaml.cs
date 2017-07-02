using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Spreadsheet;

namespace ConvertToXpsDocumentCs
{
    public partial class MainWindow : Window
    {
        XpsDocument xpsDocument;

        public MainWindow()
        {
            InitializeComponent();

            SetDocumentViewer(this.DocumentViewer);
        }

        private void SetDocumentViewer(DocumentViewer documentViewer)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            ExcelFile ef = new ExcelFile();

            var ws = ef.Worksheets.Add("Sheet1");

            ws.Cells[0, 0].Value = "English:";
            ws.Cells[0, 1].Value = "Hello";

            ws.Cells[1, 0].Value = "Russian:";
            ws.Cells[1, 1].Value = new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' });

            ws.Cells[2, 0].Value = "Chinese:";
            ws.Cells[2, 1].Value = new string(new char[] { '\u4f60', '\u597d' });

            ws.Cells[4, 0].Value = "In order to see Russian and Chinese characters you need to have appropriate fonts on your PC.";
            ws.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = true;

            ws.HeadersFooters.DefaultPage.Header.CenterSection.Content = "Export To XpsDocument / DocumentViewer Control Sample";

            ws.PrintOptions.PrintGridlines = true;

            // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
            // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
            this.xpsDocument = ef.ConvertToXpsDocument(SaveOptions.XpsDefault);

            documentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();
        }
    }
}
