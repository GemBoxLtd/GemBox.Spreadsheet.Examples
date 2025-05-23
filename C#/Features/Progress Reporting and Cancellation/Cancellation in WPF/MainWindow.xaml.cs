using GemBox.Spreadsheet;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

public partial class MainWindow : Window
{
    private volatile bool cancellationRequested;

    public MainWindow()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private async void loadButton_Click(object sender, RoutedEventArgs e)
    {
        // Capture the current context on the UI thread.
        var context = SynchronizationContext.Current;

        var loadOptions = new XlsxLoadOptions();
        loadOptions.ProgressChanged += (eventSender, args) =>
        {
            // Show progress.
            context.Post(progressPercentage => this.progressBar.Value = (int)progressPercentage, args.ProgressPercentage);

            // Cancel if requested.
            if (this.cancellationRequested)
                args.CancelOperation();
        };

        try
        {
            var file = await Task.Run(() => ExcelFile.Load("LargeFile.xlsx", loadOptions));
        }
        catch (OperationCanceledException)
        {
            // Operation cancelled.
        }
    }

    private void cancelButton_Click(object sender, RoutedEventArgs e)
    {
        this.cancellationRequested = true;
    }
}
