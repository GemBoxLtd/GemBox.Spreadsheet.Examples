using GemBox.Spreadsheet;
using System.Collections.Generic;

namespace BlazorServerApp.Data
{
    public class ReportModel
    {
        public IList<ReportItemModel> Items { get; } = new List<ReportItemModel>()
        {
            new ReportItemModel() { Id = 100, Name = "John Doe", Salary = 3600 },
            new ReportItemModel() { Id = 101, Name = "Jane Doe", Salary = 7200 },
            new ReportItemModel() { Id = 102, Name = "Fred Nurk", Salary = 2580 },
            new ReportItemModel() { Id = 103, Name = "Hans Meier", Salary = 3200 },
            new ReportItemModel() { Id = 104, Name = "Ivan Horvat", Salary = 4100 },
            new ReportItemModel() { Id = 105, Name = "Jean Dupont", Salary = 6850 },
            new ReportItemModel() { Id = 106, Name = "Mario Rossi", Salary = 4400 }
        };
        public string Format { get; set; } = "XLSX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["XLSX"] = new XlsxSaveOptions(),
            ["XLS"] = new XlsSaveOptions(),
            ["ODS"] = new OdsSaveOptions(),
            ["PDF"] = new PdfSaveOptions(),
            ["HTML"] = new HtmlSaveOptions() { EmbedImages = true },
            ["MHTML"] = new HtmlSaveOptions() { HtmlType = HtmlType.Mhtml },
            ["CSV"] = new CsvSaveOptions(CsvType.CommaDelimited),
            ["TXT"] = new CsvSaveOptions(CsvType.TabDelimited),
            ["XPS"] = new XpsSaveOptions(), // XPS is supported only on Windows.
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff),
            ["SVG"] = new ImageSaveOptions(ImageSaveFormat.Svg)
        };
    }

    public class ReportItemModel
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public int Salary { get; set; }
    }
}