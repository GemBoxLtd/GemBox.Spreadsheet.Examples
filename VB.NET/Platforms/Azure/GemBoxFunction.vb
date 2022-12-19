Imports System.IO
Imports Microsoft.AspNetCore.Mvc
Imports Microsoft.Azure.WebJobs
Imports Microsoft.Azure.WebJobs.Extensions.Http
Imports Microsoft.AspNetCore.Http
Imports Microsoft.Extensions.Logging
Imports GemBox.Spreadsheet

Module GemBoxFunction
#Disable Warning BC42356 ' This async method lacks 'Await'.
    <FunctionName("GemBoxFunction")>
    Async Function Run(<HttpTrigger(AuthorizationLevel.Anonymous, "get", Route:=Nothing)> req As HttpRequest, log As ILogger) As Task(Of IActionResult)
#Enable Warning BC42356

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Hello World")

        worksheet.Cells(0, 0).Value = "Hello"
        worksheet.Cells(0, 1).Value = "World"

        Dim fileName = "Output.xlsx"
        Dim options = SaveOptions.XlsxDefault

        Using stream As new MemoryStream()
            workbook.Save(stream, options)
            return New FileContentResult(stream.ToArray(), options.ContentType) With { .FileDownloadName = fileName }
        End Using

    End Function
End Module
