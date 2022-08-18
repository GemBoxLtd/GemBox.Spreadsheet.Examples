using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using GemBox.Spreadsheet;

public static class GemBoxFunction
{
    [FunctionName("GemBoxFunction")]
#pragma warning disable CS1998 // Async method lacks 'await' operators.
    public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req, ILogger log)
#pragma warning restore CS1998
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Hello World");

        worksheet.Cells[0, 0].Value = "Hello";
        worksheet.Cells[0, 1].Value = "World";

        var fileName = "Output.xlsx";
        var options = SaveOptions.XlsxDefault;

        using (var stream = new MemoryStream())
        {
            workbook.Save(stream, options);
            return new FileContentResult(stream.ToArray(), options.ContentType) { FileDownloadName = fileName };
        }
    }
}