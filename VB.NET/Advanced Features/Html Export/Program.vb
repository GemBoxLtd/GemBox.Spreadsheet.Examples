Imports System.Linq
Imports System.Xml
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim workbook = ExcelFile.Load("HtmlExport.xlsx")

        Dim worksheet = workbook.Worksheets(0)

        ' Set some ExcelPrintOptions properties for HTML export.
        worksheet.PrintOptions.PrintHeadings = True
        worksheet.PrintOptions.PrintGridlines = True

        ' Specify cell range which should be exported to HTML.
        worksheet.NamedRanges.SetPrintArea(worksheet.Cells.GetSubrange("A1", "J42"))

        Dim options = New HtmlSaveOptions() With
        {
            .HtmlType = HtmlType.Html,
            .SelectionType = SelectionType.EntireFile
        }

        workbook.Save("HtmlExport.html", options)
    End Sub

    Sub Example2()
        Dim workbook = ExcelFile.Load("HtmlExport.xlsx")

        ' Specify exporting of Excel data as an HTML table with embedded images.
        Dim options As New HtmlSaveOptions() With
        {
            .EmbedImages = True,
            .HtmlType = HtmlType.HtmlTable
        }

        Using writer = XmlWriter.Create("SingleHtmlExport.html",
            New XmlWriterSettings() With {.OmitXmlDeclaration = True})

            writer.WriteStartElement("html")
            writer.WriteStartElement("body")

            ' Write Excel sheets to a single HTML file in reverse order.
            For Each worksheet In workbook.Worksheets.Reverse()
                writer.WriteElementString("h1", worksheet.Name)

                workbook.Worksheets.ActiveWorksheet = worksheet
                workbook.Save(writer, options)
            Next

            writer.WriteEndDocument()
        End Using
    End Sub

End Module