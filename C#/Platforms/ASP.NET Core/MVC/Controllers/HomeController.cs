using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using SpreadsheetCoreMvc.Models;
using GemBox.Spreadsheet;

namespace SpreadsheetCoreMvc.Controllers
{
    public class HomeController : Controller
    {
        static HomeController()
        {
            // If using the Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public IActionResult Index()
        {
            return View(new ReportModel());
        }

        public FileStreamResult Download(ReportModel model)
        {
            // Create new spreadsheet.
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Report");

            // Set styles on rows and columns.
            worksheet.Rows[0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            worksheet.Rows[0].Style.Font.Weight = ExcelFont.BoldWeight;
            worksheet.Columns[0].SetWidth(40, LengthUnit.Pixel);
            worksheet.Columns[1].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[2].SetWidth(100, LengthUnit.Pixel);
            worksheet.Columns[2].Style.NumberFormat = @"\$\ #,##0";

            // Create header row.
            worksheet.Cells["A1"].Value = nameof(ReportItemModel.Id);
            worksheet.Cells["B1"].Value = nameof(ReportItemModel.Name);
            worksheet.Cells["C1"].Value = nameof(ReportItemModel.Salery);

            // Create data rows.
            for (int r = 1; r <= model.Items.Count; r++)
            {
                ReportItemModel item = model.Items[r - 1];
                worksheet.Cells[r, 0].Value = item.Id;
                worksheet.Cells[r, 1].Value = item.Name;
                worksheet.Cells[r, 2].Value = item.Salery;
            }

            // Save spreadsheet in specified file format.
            var stream = new MemoryStream();
            workbook.Save(stream, model.Options);

            // Download file.
            return File(stream, model.Options.ContentType, $"OutputFromView.{model.Format.ToLower()}");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel() { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

namespace SpreadsheetCoreMvc.Models
{
    public class ReportModel
    {
        public IList<ReportItemModel> Items { get; set; } = new List<ReportItemModel>()
        {
            new ReportItemModel() { Id = 100, Name = "John Doe", Salery = 3600 },
            new ReportItemModel() { Id = 101, Name = "Jane Doe", Salery = 7200 },
            new ReportItemModel() { Id = 102, Name = "Fred Nurk", Salery = 2580 },
            new ReportItemModel() { Id = 103, Name = "Hans Meier", Salery = 3200 },
            new ReportItemModel() { Id = 104, Name = "Ivan Horvat", Salery = 4100 },
            new ReportItemModel() { Id = 105, Name = "Jean Dupont", Salery = 6850 },
            new ReportItemModel() { Id = 106, Name = "Mario Rossi", Salery = 4400 }
        };
        public string Format { get; set; } = "XLSX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["XLSX"] = new XlsxSaveOptions(),
            ["XLS"] = new XlsSaveOptions(),
            ["ODS"] = new OdsSaveOptions(),
            ["CSV"] = new CsvSaveOptions(CsvType.CommaDelimited),
            ["PDF"] = new PdfSaveOptions(),
            ["HTML"] = new HtmlSaveOptions() { EmbedImages = true },
            ["XPS"] = new XpsSaveOptions(),
            ["BMP"] = new ImageSaveOptions(ImageSaveFormat.Bmp),
            ["PNG"] = new ImageSaveOptions(ImageSaveFormat.Png),
            ["JPG"] = new ImageSaveOptions(ImageSaveFormat.Jpeg),
            ["GIF"] = new ImageSaveOptions(ImageSaveFormat.Gif),
            ["TIF"] = new ImageSaveOptions(ImageSaveFormat.Tiff)
        };
    }

    public class ReportItemModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Salery { get; set; }
    }
}
