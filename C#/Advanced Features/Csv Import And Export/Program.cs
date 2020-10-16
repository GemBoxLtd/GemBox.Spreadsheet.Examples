using System;
using System.IO;
using System.Linq;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
        Example3();
        Example4();
        Example5();
    }

    static void Example1()
    {
        // Read CSV file.
        var workbook = ExcelFile.Load("Input.csv", new CsvLoadOptions(CsvType.CommaDelimited));

        // Add new row.
        var worksheet = workbook.Worksheets[0];
        var row = worksheet.Rows[worksheet.Rows.Count];
        row.Cells[0].Value = "Jane Doe";
        row.Cells[1].Value = 3500;
        row.Cells[2].Value = 35;

        // Write CSV file.
        workbook.Save("Output.csv", new CsvSaveOptions(CsvType.CommaDelimited));
    }

    static void Example2()
    {
        var csvOptions = new CsvLoadOptions(CsvType.CommaDelimited)
        {
            AllowNewLineInQuotes = true,
            HasQuotedValues = true,
            HasFormulas = true
        };

        // Read CSV file using specified CsvLoadOptions.
        var workbook = ExcelFile.Load("ArtificalObjectsOnMoon.csv", csvOptions);

        // Calculate Excel formulas from CSV data.
        var worksheet = workbook.Worksheets[0];
        worksheet.Calculate();

        // Iterate through read CSV records.
        foreach (var row in worksheet.Rows)
        {
            // Iterate through read CSV fields.
            foreach (var cell in row.AllocatedCells)
            {
                // Display just the first line of text from Excel cell.
                var value = cell.Value?.ToString() ?? string.Empty;
                Console.Write($"{value.Split('\n')[0],-25}");
            }

            Console.WriteLine();
        }
    }

    static void Example3()
    {
        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        // Tabular sample data for exporting into a CSV file.
        var skyscrapers = new object[,]
        {
            { "Rank", "Building", "City", "Country", "Height (m)", "Height (ft)", "Floors", "Built" },
            { 1, "Burj Khalifa", "Dubai", "United Arab Emirates", 829.8, 2722, 163, 2010 },
            { 2, "Shanghai Tower", "Shanghai", "China", 632, 2073, 128, 2015 },
            { 3, "Abraj Al-Bait Towers", "Mecca", "Saudi Arabia", 601, 1971, 120, 2012 },
            { 4, "Ping An Finance Center", "Shenzhen", "China", 599, 1965, 115, 2016 },
            { 5, "Lotte World Tower", "Seoul", "South Korea", 555.7, 1823, 123, 2016 },
            { 6, "One World Trade Center", "New York City", "United States", 546.2, 1792, 104, 2014 },
            { 7, "Guangzhou CTF Finance Centre", "Guangzhou", "China", 530, 1739, 111, 2016 },
            { 7, "Tianjin CTF Finance Centre", "Tianjin", "China", 530, 1739, 98, 2018 },
            { 9, "China Zun", "Beijing", "China", 528, 1732, 108, 2018 },
            { 10, "Willis Tower", "Chicago", "United States", 527, 1729, 108, 1974 },
            { 11, "Taipei 101", "Taipei", "Taiwan", 508, 1667, 101, 2004 },
            { 12, "Shanghai World Financial Center", "Shanghai", "China", 494.3, 1622, 101, 2008 },
            { 13, "International Commerce Centre", "Hong Kong", "China", 484, 1588, 118, 2010 },
            { 15, "Central Park Tower", "New York City", "United States", 472.4, 1550, 103, 2020 },
            { 16, "Landmark 81", "Ho Chi Minh City", "Vietnam", 469.5, 1540, 81, 2018 },
            { 17, "Lakhta Center", "St. Petersburg", "Russia", 462, 1516, 86, 2018 },
            { 18, "John Hancock Center", "Chicago", "United States", 456.9, 1499, 100, 1969 },
            { 19, "Changsha IFS Tower T1", "Changsha", "China", 452, 1483, 94, 2017 },
            { 20, "Petronas Tower 1", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998 },
            { 20, "Petronas Tower 2", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998 },
            { 22, "Zifeng Tower", "Nanjing", "China", 450, 1476, 89, 2009 },
            { 22, "Suzhou IFS", "Suzhou", "China", 450, 1476, 98, 2017 },
            { 24, "The Exchange 106", "Kuala Lumpur", "Malaysia", 445.1, 1460, 95, 2018 },
            { 25, "Empire State Building", "New York City", "United States", 443.2, 1454, 102, 1931 },
            { 26, "Kingkey 100", "Shenzhen", "China", 442, 1449, 100, 2011 },
            { 27, "Guangzhou International Finance Center", "Guangzhou", "China", 438.6, 1445, 103, 2009 },
            { 28, "Wuhan Center", "Wuhan", "China", 438, 1437, 88, 2017 },
            { 29, "111 West 57th Street", "New York City", "United States", 435.3, 1428, 82, 2019 },
            { 30, "Dongguan International Trade Center 1", "Dongguan", "China", 426.9, 1401, 88, 2019 },
            { 31, "One Vanderbilt", "New York City", "United States", 427, 1401, 58, 2019 },
            { 32, "432 Park Avenue", "New York City", "United States", 425.5, 1396, 85, 2015 },
            { 33, "Marina 101", "Dubai", "United Arab Emirates", 425, 1394, 101, 2017 },
            { 34, "Trump International Hotel and Tower", "Chicago", "United States", 423.2, 1388, 96, 2009 },
            { 35, "Jin Mao Tower", "Shanghai", "China", 421, 1381, 88, 1998 },
            { 36, "Princess Tower", "Dubai", "United Arab Emirates", 414, 1358, 101, 2012 },
            { 37, "Al Hamra Tower", "Kuwait City", "Kuwait", 412.6, 1354, 80, 2010 },
            { 38, "Two International Finance Centre", "Hong Kong", "China", 412, 1352, 88, 2003 },
            { 39, "Haeundae LCT The Sharp Landmark Tower", "Busan", "South Korea", 411.6, 1350, 101, 2019 },
            { 40, "Guangxi China Resources Tower", "Nanning", "China", 402.7, 1321, 85, 2018 },
            { 41, "Guiyang Financial Center Tower 1", "Guiyang", "China", 401, 1316, 79, 2020 }
        };

        // Write data into Excel cells.
        int rowCount = skyscrapers.GetLength(0);
        int columnCount = skyscrapers.GetLength(1);
        for (int row = 0; row < rowCount; row++)
            for (int column = 0; column < columnCount; column++)
                worksheet.Cells[row, column].Value = skyscrapers[row, column];

        // Format Excel columns.
        worksheet.Columns["E"].Style.NumberFormat = @"0.0 \m";
        worksheet.Columns["F"].Style.NumberFormat = @"0,000 \f\t";

        var csvOptions = new CsvSaveOptions(CsvType.CommaDelimited)
        {
            UseFormattedValues = true
        };

        // Write CSV file using specified CsvSaveOptions.
        workbook.Save("Skyscrapers.csv", csvOptions);
    }

    static void Example4()
    {
        // Create large CSV file.
        using (var csv = File.CreateText("large-file.csv"))
            for (int i = 0; i < 5_000_000; i++)
                csv.WriteLine(i);

        // Import all CSV data into multiple sheets.
        var workbook = LargeCsvReader.ReadFile("large-file.csv", LoadOptions.CsvDefault);

        // Display name and rows count of generated sheets.
        foreach (var worksheet in workbook.Worksheets)
            Console.WriteLine($"Name: {worksheet.Name} | Rows: {worksheet.Rows.Count:#,###}");
    }

    static void Example5()
    {
        // Create large ExcelFile.
        ExcelFile workbook = new ExcelFile();
        ExcelWorksheet worksheet = null;

        int max = 1_048_576;
        for (int index = 0; index < 5_000_000; index++)
        {
            int current = index % max;
            if (current == 0)
                worksheet = workbook.Worksheets.Add($"Sheet{index / max}");
            worksheet.Cells[current, 0].SetValue(index);
        }

        // Export multiple sheets into single CSV file.
        var options = SaveOptions.CsvDefault;
        using (var writer = File.CreateText("large-file.csv"))
            foreach (var sheet in workbook.Worksheets)
            {
                workbook.Worksheets.ActiveWorksheet = sheet;
                workbook.Save(writer, options);
            }

        // Display number of lines, or records, in generated CSV file.
        int csvLinesCount = File.ReadLines("large-file.csv").Count();
        Console.WriteLine($"Records: {csvLinesCount:#,###}");
    }
}