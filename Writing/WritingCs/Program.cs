using System.Drawing;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Writing");

        // Tabular sample data for writing into an Excel file.
        object[,] skyscrapers = new object[21, 7]
        {
            {"Rank", "Building", "City", "Metric", "Imperial", "Floors", "Built (Year)"},
            { 1, "Taipei 101", "Taipei", 509, 1671, 101, 2004},
            { 2, "Petronas Tower 1", "Kuala Lumpur", 452, 1483, 88, 1998},
            { 3, "Petronas Tower 2", "Kuala Lumpur", 452, 1483, 88, 1998},
            { 4, "Sears Tower", "Chicago", 442, 1450, 108, 1974},
            { 5, "Jin Mao Tower", "Shanghai", 421, 1380, 88, 1998},
            { 6, "2 International Finance Centre", "Hong Kong", 415, 1362, 88, 2003},
            { 7, "CITIC Plaza", "Guangzhou", 391, 1283, 80, 1997},
            { 8, "Shun Hing Square", "Shenzhen", 384, 1260, 69, 1996},
            { 9, "Empire State Building", "New York City", 381, 1250, 102, 1931},
            {10, "Central Plaza", "Hong Kong", 374, 1227, 78, 1992},
            {11, "Bank of China Tower", "Hong Kong", 367, 1205, 72, 1990},
            {12, "Emirates Office Tower", "Dubai", 355, 1163, 54, 2000},
            {13, "Tuntex Sky Tower", "Kaohsiung", 348, 1140, 85, 1997},
            {14, "Aon Center", "Chicago", 346, 1136, 83, 1973},
            {15, "The Center", "Hong Kong", 346, 1135, 73, 1998},
            {16, "John Hancock Center", "Chicago", 344, 1127, 100, 1969},
            {17, "Ryugyong Hotel", "Pyongyang", 330, 1083, 105, 1992},
            {18, "Burj Al Arab", "Dubai", 321, 1053, 60, 1999},
            {19, "Chrysler Building", "New York City", 319, 1046, 77, 1930},
            {20, "Bank of America Plaza", "Atlanta", 312, 1023, 55, 1992}
        };

        ws.Cells[0, 0].Value = "Example of writing typical table - tallest buildings in the world (2004):";

        // Column width of 8, 30, 16, 9, 9, 9, 9, 4 and 5 characters.
        ws.Columns[0].Width = 8 * 256;
        ws.Columns[1].Width = 30 * 256;
        ws.Columns[2].Width = 16 * 256;
        ws.Columns[3].Width = 9 * 256;
        ws.Columns[4].Width = 9 * 256;
        ws.Columns[5].Width = 9 * 256;
        ws.Columns[6].Width = 9 * 256;
        ws.Columns[7].Width = 4 * 256;
        ws.Columns[8].Width = 5 * 256;

        int i, j;
        // Write header data to Excel cells.
        for (j = 0; j < 7; j++)
            ws.Cells[3, j].Value = skyscrapers[0, j];

        ws.Cells.GetSubrangeAbsolute(2, 0, 3, 0).Merged = true;
        ws.Cells.GetSubrangeAbsolute(2, 1, 3, 1).Merged = true;
        ws.Cells.GetSubrangeAbsolute(2, 2, 3, 2).Merged = true;
        ws.Cells.GetSubrangeAbsolute(2, 3, 2, 4).Merged = true;
        ws.Cells[2, 3].Value = "Height";
        ws.Cells.GetSubrangeAbsolute(2, 5, 3, 5).Merged = true;
        ws.Cells.GetSubrangeAbsolute(2, 6, 3, 6).Merged = true;

        CellStyle tmpStyle = new CellStyle();
        tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
        tmpStyle.FillPattern.SetSolid(Color.Chocolate);
        tmpStyle.Font.Weight = ExcelFont.BoldWeight;
        tmpStyle.Font.Color = Color.White;
        tmpStyle.WrapText = true;
        tmpStyle.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top, Color.Black, LineStyle.Thin);

        ws.Cells.GetSubrangeAbsolute(2, 0, 3, 6).Style = tmpStyle;

        tmpStyle = new CellStyle();
        tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
        tmpStyle.Font.Weight = ExcelFont.BoldWeight;

        CellRange mergedRange = ws.Cells.GetSubrangeAbsolute(4, 7, 13, 7);
        mergedRange.Merged = true;
        mergedRange.Value = "T o p   1 0";
        tmpStyle.Rotation = -90;
        tmpStyle.FillPattern.SetSolid(Color.Lime);
        mergedRange.Style = tmpStyle;

        mergedRange = ws.Cells.GetSubrangeAbsolute(4, 8, 23, 8);
        mergedRange.Merged = true;
        mergedRange.Value = "T o p   2 0";
        tmpStyle.IsTextVertical = true;
        tmpStyle.FillPattern.SetSolid(Color.Gold);
        mergedRange.Style = tmpStyle;

        mergedRange = ws.Cells.GetSubrangeAbsolute(14, 7, 23, 7);
        mergedRange.Merged = true;
        mergedRange.Style = tmpStyle;

        // Write and format sample data to Excel cells.
        for (i = 0; i < 20; i++)
            for (j = 0; j < 7; j++)
            {
                ExcelCell cell = ws.Cells[i + 4, j];

                cell.Value = skyscrapers[i + 1, j];

                if (i % 2 == 0)
                    cell.Style.FillPattern.SetSolid(Color.LightSkyBlue);
                else
                    cell.Style.FillPattern.SetSolid(Color.FromArgb(210, 210, 230));

                if (j == 3)
                    cell.Style.NumberFormat = "#\" m\"";

                if (j == 4)
                    cell.Style.NumberFormat = "#\" ft\"";

                if (j > 2)
                    cell.Style.Font.Name = "Courier New";

                cell.Style.Borders[IndividualBorder.Right].LineStyle = LineStyle.Thin;
            }

        ws.Cells.GetSubrange("A5", "I24").Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Double);
        ws.Cells.GetSubrange("A3", "G4").Style.Borders.SetBorders(MultipleBorders.Vertical | MultipleBorders.Top, Color.Black, LineStyle.Double);
        ws.Cells.GetSubrange("A5", "H14").Style.Borders.SetBorders(MultipleBorders.Bottom | MultipleBorders.Right, Color.Black, LineStyle.Double);

        ws.Cells["A27"].Value = "Notes:";
        ws.Cells["A28"].Value = "a) \"Metric\" and \"Imperial\" columns use custom number formatting.";
        ws.Cells["A29"].Value = "b) All number columns use \"Courier New\" font for improved number readability.";
        ws.Cells["A30"].Value = "c) Multiple merged ranges were used for table header and categories header.";

        ws.PrintOptions.FitWorksheetWidthToPages = 1;

        ef.Save("Writing.xlsx");
    }
}
