using GemBox.Spreadsheet;
using System.Windows;
using System.Threading;
using System.Threading.Tasks;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        InitializeComponent();
    }

    private async void loadButton_Click(object sender, RoutedEventArgs e)
    {
        // Capture the current context on UI thread.
        var context = SynchronizationContext.Current;

        // Create load options.
        var loadOptions = new XlsxLoadOptions();
        loadOptions.ProgressChanged += (eventSender, args) =>
        {
            var percentage = args.ProgressPercentage;
            // Invoke on UI thread.
            context.Post(progressPercentage =>
            {
                // Update UI.
                this.progressBar.Value = (int)progressPercentage;
                this.percentageLabel.Content = progressPercentage.ToString() + "%";
            }, percentage);
        };

        this.percentageLabel.Content = "0%";
        // Use tasks to run the load operation in a new thread.
        var file = await Task.Run(() => ExcelFile.Load("LargeFile.xlsx", loadOptions));
    }
}
