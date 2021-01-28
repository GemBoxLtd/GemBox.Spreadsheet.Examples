using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;
using GemBox.Spreadsheet;

public partial class Form1 : Form
{
    private ExcelFile workbook;

    public Form1()
    {
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private void LoadFileMenuItem_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter =
            "XLSX files (*.xlsx, *.xltx, *.xlsm, *.xltm)|*.xlsx;*.xltx;*.xlsm;*.xltm" +
            "|XLS files (*.xls, *.xlt)|*.xls;*.xlt" +
            "|ODS files (*.ods, *.ots)|*.ods;*.ots" +
            "|CSV files (*.csv, *.tsv)|*.csv;*.tsv" +
            "|HTML files (*.html, *.htm)|*.html;*.htm";

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            this.workbook = ExcelFile.Load(openFileDialog.FileName);
            this.ShowPrintPreview();
        }
    }

    private void PrintFileMenuItem_Click(object sender, EventArgs e)
    {
        if (this.workbook == null)
            return;

        PrintDialog printDialog = new PrintDialog() { AllowSomePages = true };
        if (printDialog.ShowDialog() == DialogResult.OK)
        {
            PrinterSettings printerSettings = printDialog.PrinterSettings;
            PrintOptions printOptions = new PrintOptions() { SelectionType = SelectionType.EntireFile };

            // Set PrintOptions properties based on PrinterSettings properties.
            printOptions.CopyCount = printerSettings.Copies;
            printOptions.FromPage = printerSettings.FromPage == 0 ? 0 : printerSettings.FromPage - 1;
            printOptions.ToPage = printerSettings.ToPage == 0 ? int.MaxValue : printerSettings.ToPage - 1;

            this.workbook.Print(printerSettings.PrinterName, printOptions);
        }
    }

    private void ShowPrintPreview()
    {
        // Create image for each Excel workbook's page.
        Image[] images = this.CreatePrintPreviewImages();
        int imageIndex = 0;

        // Draw each page's image on PrintDocument for print preview.
        var printDocument = new PrintDocument();
        printDocument.PrintPage += (sender, e) =>
        {
            using (Image image = images[imageIndex])
            {
                var graphics = e.Graphics;
                var region = graphics.VisibleClipBounds;

                // Rotate image if it has landscape orientation.
                if (image.Width > image.Height)
                    image.RotateFlip(RotateFlipType.Rotate270FlipNone);

                graphics.DrawImage(image, 0, 0, region.Width, region.Height);
            }

            ++imageIndex;
            e.HasMorePages = imageIndex < images.Length;
        };

        this.PageUpDown.Value = 1;
        this.PageUpDown.Maximum = images.Length;
        this.PrintPreviewControl.Document = printDocument;
    }

    private Image[] CreatePrintPreviewImages()
    {
        var paginatorOptions = new PaginatorOptions { SelectionType = SelectionType.EntireFile };
        var pages = this.workbook.GetPaginator(paginatorOptions).Pages;

        var images = new Image[pages.Count];
        var imageOptions = new ImageSaveOptions();

        for (int pageIndex = 0; pageIndex < pages.Count; ++pageIndex)
        {
            var imageStream = new MemoryStream();
            pages[pageIndex].Save(imageStream, imageOptions);
            images[pageIndex] = Image.FromStream(imageStream);
        }

        return images;
    }

    private void PageUpDown_ValueChanged(object sender, EventArgs e)
    {
        this.PrintPreviewControl.StartPage = (int)this.PageUpDown.Value - 1;
    }
}