using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using GemBox.Spreadsheet;

public class GemBoxFunction
{
    [Function("GemBoxFunction")]
    public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Hello World");

        worksheet.Cells[0, 0].Value = "Hello";
        worksheet.Cells[0, 1].Value = "World";

        var fileName = "Output.xlsx";
        var options = SaveOptions.XlsxDefault;

        using var stream = new MemoryStream();
        workbook.Save(stream, options);
        var bytes = stream.ToArray();

        var response = req.CreateResponse(HttpStatusCode.OK);
        response.Headers.Add("Content-Type", options.ContentType);
        response.Headers.Add("Content-Disposition", "attachment; filename=" + fileName);
        await response.Body.WriteAsync(bytes, 0, bytes.Length);
        return response;
    }
}