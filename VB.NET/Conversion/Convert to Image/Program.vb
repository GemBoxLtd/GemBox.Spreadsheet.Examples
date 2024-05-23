Imports GemBox.Spreadsheet
Imports System.IO
Imports System.IO.Compression

Module Program

    Sub Main()

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load an Excel file into the ExcelFile object.
        Dim workbook = ExcelFile.Load("CombinedTemplate.xlsx")

        ' Create image save options.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Png) With
        {
            .PageNumber = 0, ' Select the first Excel page.
            .Width = 1240, ' Set the image width.
            .CropToContent = True ' Export only the sheet's content.
        }

        ' Save the ExcelFile object to a PNG file.
        workbook.Save("Output.png", imageOptions)
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load an Excel file.
        Dim workbook = ExcelFile.Load("CombinedTemplate.xlsx")

        ' Max integer value indicates that all spreadsheet pages should be saved.
        Dim imageOptions As New ImageSaveOptions(ImageSaveFormat.Tiff) With
        {
            .SelectionType = SelectionType.EntireFile,
            .PageCount = Integer.MaxValue
        }

        ' Save the TIFF file with multiple frames, each frame represents a single Excel page.
        workbook.Save("Output.tiff", imageOptions)
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load an Excel file.
        Dim workbook = ExcelFile.Load("CombinedTemplate.xlsx")

        ' Get Excel pages.
        Dim paginatorOptions As New PaginatorOptions() With {.SelectionType = SelectionType.EntireFile}
        Dim pages = workbook.GetPaginator(paginatorOptions).Pages

        ' Create a ZIP file for storing PNG files.
        Using archiveStream = File.OpenWrite("Output.zip")
            Using archive As New ZipArchive(archiveStream, ZipArchiveMode.Create)

                Dim imageOptions As New ImageSaveOptions()

                ' Iterate through the Excel pages.
                For pageIndex As Integer = 0 To pages.Count - 1

                    Dim page As ExcelFilePage = pages(pageIndex)

                    ' Create a ZIP entry for each spreadsheet page.
                    Dim entry = archive.CreateEntry($"Page {pageIndex + 1}.png")

                    ' Save each spreadsheet page as a PNG image to the ZIP entry.
                    Using imageStream As New MemoryStream()
                        Using entryStream = entry.Open()
                            page.Save(imageStream, imageOptions)
                            imageStream.CopyTo(entryStream)
                        End Using
                    End Using
                Next

            End Using
        End Using
    End Sub

End Module
