using GemBox.Spreadsheet;

namespace SpreadsheetMaui
{
    public partial class MainPage : ContentPage
    {
        static MainPage()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public MainPage()
        {
            InitializeComponent();
        }

        private async Task<string> CreateWorkbookAsync()
        {
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            foreach (var cell in table.Root[0].Cast<EntryCell>())
                worksheet.Cells[cell.Label].Value = cell.Text;

            worksheet.Columns["A"].AutoFit();

            using var stream = new MemoryStream();
            using (var imageStream = await FileSystem.OpenAppPackageFileAsync("dices.png"))
                await imageStream.CopyToAsync(stream);
            worksheet.Pictures.Add(stream, ExcelPictureFormat.Png, "C1", "E5");

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Example.pdf");

            workbook.Save(filePath);

            return filePath;
        }

        private async void Button_Clicked(object sender, EventArgs e)
        {
            button.IsEnabled = false;
            activity.IsRunning = true;

            var filePath = await CreateWorkbookAsync();
            await Launcher.OpenAsync(new OpenFileRequest(Path.GetFileName(filePath), new ReadOnlyFile(filePath)));

            activity.IsRunning = false;
            button.IsEnabled = true;
        }
    }
}