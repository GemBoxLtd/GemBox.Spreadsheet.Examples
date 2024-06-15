Imports GemBox.Spreadsheet
Imports Microsoft.Azure.Functions.Worker
Imports Microsoft.Azure.Functions.Worker.Http
Imports System.IO
Imports System.Net
Imports System.Threading.Tasks

Public Class GemBoxFunction
    <[Function]("GemBoxFunction")>
    Public Async Function Run(<HttpTrigger(AuthorizationLevel.Anonymous, "get")> req As HttpRequestData) As Task(Of HttpResponseData)

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Hello World")

        worksheet.Cells(0, 0).Value = "Hello"
        worksheet.Cells(0, 1).Value = "World"

        Dim fileName = "Output.xlsx"
        Dim options = SaveOptions.XlsxDefault

        Using stream As New MemoryStream()
            workbook.Save(stream, options)
            Dim bytes = stream.ToArray()

            Dim response = req.CreateResponse(HttpStatusCode.OK)
            response.Headers.Add("Content-Type", options.ContentType)
            response.Headers.Add("Content-Disposition", "attachment; filename=" & fileName)
            Await response.Body.WriteAsync(bytes, 0, bytes.Length)
            Return response
        End Using

    End Function
End Class
