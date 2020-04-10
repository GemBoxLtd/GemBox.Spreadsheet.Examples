using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using GemBox.Spreadsheet;

namespace SpreadsheetCore.Controllers
{
    public class WorkbookController : Controller
    {
        private static readonly IList<WorkbookItemModel> data = new List<WorkbookItemModel>()
        {
            new WorkbookItemModel() { Id = 100, FirstName = "John", LastName = "Doe"},
            new WorkbookItemModel() { Id = 101, FirstName = "Fred", LastName = "Nurk"},
            new WorkbookItemModel() { Id = 102, FirstName = "Hans", LastName = "Meier"},
            new WorkbookItemModel() { Id = 103, FirstName = "Ivan", LastName = "Horvat"},
            new WorkbookItemModel() { Id = 104, FirstName = "Jean", LastName = "Dupont"},
            new WorkbookItemModel() { Id = 105, FirstName = "Mario", LastName = "Rossi"},
        };

        static WorkbookController()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View(new WorkbookModel()
            {
                Items = data,
                SelectedFormat = "XLSX"
            });
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult Create(WorkbookModel model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var book = new ExcelFile();
            var sheet = book.Worksheets.Add("Sheet1");

            CellStyle style = sheet.Rows[0].Style;
            style.Font.Weight = ExcelFont.BoldWeight;
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Columns[0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            sheet.Columns[0].SetWidth(50, LengthUnit.Pixel);
            sheet.Columns[1].SetWidth(150, LengthUnit.Pixel);
            sheet.Columns[2].SetWidth(150, LengthUnit.Pixel);

            sheet.Cells["A1"].Value = "ID";
            sheet.Cells["B1"].Value = "First Name";
            sheet.Cells["C1"].Value = "Last Name";

            for (int r = 1; r <= model.Items.Count; r++)
            {
                WorkbookItemModel item = model.Items[r - 1];
                sheet.Cells[r, 0].Value = item.Id;
                sheet.Cells[r, 1].Value = item.FirstName;
                sheet.Cells[r, 2].Value = item.LastName;
            }

            SaveOptions options = GetSaveOptions(model.SelectedFormat);

            using (var stream = new MemoryStream())
            {
                book.Save(stream, options);
                return File(stream.ToArray(), options.ContentType, "Create." + model.SelectedFormat.ToLower());
            }
        }

        private static SaveOptions GetSaveOptions(string format)
        {
            switch (format.ToUpper())
            {
                case "XLSX":
                    return SaveOptions.XlsxDefault;
                case "XLS":
                    return SaveOptions.XlsDefault;
                case "ODS":
                    return SaveOptions.OdsDefault;
                case "CSV":
                    return SaveOptions.CsvDefault;
                case "HTML":
                    return SaveOptions.HtmlDefault;
                case "PDF":
                    return SaveOptions.PdfDefault;

                case "XPS":
                case "PNG":
                case "JPG":
                case "GIF":
                case "TIF":
                case "BMP":
                case "WMP":
                    throw new InvalidOperationException("To enable saving to XPS or image format, add 'Microsoft.WindowsDesktop.App' framework reference.");

                default:
                    throw new NotSupportedException();
            }
        }
    }

    public class WorkbookModel
    {
        public string SelectedFormat { get; set; }
        public IList<WorkbookItemModel> Items { get; set; }
    }

    public class WorkbookItemModel
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }
}