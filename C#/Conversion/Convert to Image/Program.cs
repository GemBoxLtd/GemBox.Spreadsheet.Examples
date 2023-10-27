using System.IO;
using System.IO.Compression;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // Load an Excel file into the ExcelFile object.
        var workbook = ExcelFile.Load("CombinedTemplate.xlsx");

        // Create image save options.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png)
        {
            PageNumber = 0, // Select the first Excel page.
            Width = 1240, // Set the image width.
            CropToContent = true // Export only the sheet's content.
        };

        // Save the ExcelFile object to a PNG file.
        workbook.Save("Output.png", imageOptions);
    }

    static void Example2()
    {
        // Load an Excel file.
        var workbook = ExcelFile.Load("CombinedTemplate.xlsx");

        // Max integer value indicates that all spreadsheet pages should be saved.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Tiff)
        {
            SelectionType = SelectionType.EntireFile,
            PageCount = int.MaxValue
        };

        // Save the TIFF file with multiple frames, each frame represents a single Excel page.
        workbook.Save("Output.tiff", imageOptions);
    }

    static void Example3()
    {
        // Load an Excel file.
        var workbook = ExcelFile.Load("CombinedTemplate.xlsx");

        // Get Excel pages.
        var paginatorOptions = new PaginatorOptions() { SelectionType = SelectionType.EntireFile };
        var pages = workbook.GetPaginator(paginatorOptions).Pages;

        // Create a ZIP file for storing PNG files.
        using (var archiveStream = File.OpenWrite("Output.zip"))
        using (var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create))
        {
            var imageOptions = new ImageSaveOptions();

            // Iterate through the Excel pages.
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                ExcelFilePage page = pages[pageIndex];

                // Create a ZIP entry for each spreadsheet page.
                var entry = archive.CreateEntry($"Page {pageIndex + 1}.png");

                // Save each spreadsheet page as a PNG image to the ZIP entry.
                using (var imageStream = new MemoryStream())
                using (var entryStream = entry.Open())
                {
                    page.Save(imageStream, imageOptions);
                    imageStream.CopyTo(entryStream);
                }
            }
        }
    }
}