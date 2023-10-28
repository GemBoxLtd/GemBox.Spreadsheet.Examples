Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' Create new empty workbook.
        Dim workbook As New ExcelFile()

        ' Add new sheet.
        Dim worksheet = workbook.Worksheets.Add("Skyscrapers")

        ' Write title to Excel cell.
        worksheet.Cells("A1").Value = "List of tallest buildings (2021):"

        ' Tabular sample data for writing into an Excel file.
        Dim skyscrapers = New Object(,) _
        {
            {"Rank", "Building", "City", "Country", "Metric", "Imperial", "Floors", "Built (Year)"},
            {1, "Burj Khalifa", "Dubai", "United Arab Emirates", 828, 2717, 163, 2010},
            {2, "Shanghai Tower", "Shanghai", "China", 632, 2073, 128, 2015},
            {3, "Abraj Al-Bait Clock Tower", "Mecca", "Saudi Arabia", 601, 1971, 120, 2012},
            {4, "Ping An Finance Centre", "Shenzhen", "China", 599, 1965, 115, 2017},
            {5, "Lotte World Tower", "Seoul", "South Korea", 554.5, 1819, 123, 2016},
            {6, "One World Trade Center", "New York City", "United States", 541.3, 1776, 104, 2014},
            {7, "Guangzhou CTF Finance Centre", "Guangzhou", "China", 530, 1739, 111, 2016},
            {7, "Tianjin CTF Finance Centre", "Tianjin", "China", 530, 1739, 98, 2019},
            {9, "China Zun", "Beijing", "China", 528, 1732, 108, 2018},
            {10, "Taipei 101", "Taipei", "Taiwan", 508, 1667, 101, 2004},
            {11, "Shanghai World Financial Center", "Shanghai", "China", 492, 1614, 101, 2008},
            {12, "International Commerce Centre", "Hong Kong", "China", 484, 1588, 118, 2010},
            {13, "Central Park Tower", "New York City", "United States", 472, 1550, 98, 2020},
            {14, "Lakhta Center", "St. Petersburg", "Russia", 462, 1516, 86, 2019},
            {15, "Landmark 81", "Ho Chi Minh City", "Vietnam", 461.2, 1513, 81, 2018},
            {16, "Changsha IFS Tower T1", "Changsha", "China", 452.1, 1483, 88, 2018},
            {17, "Petronas Tower 1", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
            {17, "Petronas Tower 2", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
            {19, "Zifeng Tower", "Nanjing", "China", 450, 1476, 89, 2010},
            {19, "Suzhou IFS", "Suzhou", "China", 450, 1476, 98, 2019}
        }

        ' Set row formatting.
        worksheet.Rows("1").Style = workbook.Styles(BuiltInCellStyleName.Heading1)

        ' Set columns width.
        worksheet.Columns("A").SetWidth(8, LengthUnit.CharacterWidth)  ' Rank
        worksheet.Columns("B").SetWidth(30, LengthUnit.CharacterWidth) ' Building
        worksheet.Columns("C").SetWidth(16, LengthUnit.CharacterWidth) ' City
        worksheet.Columns("D").SetWidth(20, LengthUnit.CharacterWidth) ' Country
        worksheet.Columns("E").SetWidth(9, LengthUnit.CharacterWidth)  ' Metric
        worksheet.Columns("F").SetWidth(11, LengthUnit.CharacterWidth) ' Imperial
        worksheet.Columns("G").SetWidth(9, LengthUnit.CharacterWidth)  ' Floors
        worksheet.Columns("H").SetWidth(9, LengthUnit.CharacterWidth)  ' Built (Year)
        worksheet.Columns("I").SetWidth(4, LengthUnit.CharacterWidth)  ' Top 10
        worksheet.Columns("J").SetWidth(5, LengthUnit.CharacterWidth)  ' Top 20

        ' Write header data to Excel cells.
        For col As Integer = 0 To skyscrapers.GetLength(1) - 1
            worksheet.Cells(3, col).Value = skyscrapers(0, col)
        Next
        worksheet.Cells("E3").Value = "Height"

        worksheet.Cells.GetSubrange("A3:A4").Merged = True  ' Rank
        worksheet.Cells.GetSubrange("B3:B4").Merged = True  ' Building
        worksheet.Cells.GetSubrange("C3:C4").Merged = True  ' City
        worksheet.Cells.GetSubrange("D3:D4").Merged = True  ' Country
        worksheet.Cells.GetSubrange("E3:F3").Merged = True  ' Height
        worksheet.Cells.GetSubrange("G3:G4").Merged = True  ' Floors
        worksheet.Cells.GetSubrange("H3:H4").Merged = True  ' Built (Year)

        ' Set header cells formatting.
        Dim style = New CellStyle()
        style.HorizontalAlignment = HorizontalAlignmentStyle.Center
        style.VerticalAlignment = VerticalAlignmentStyle.Center
        style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(237, 125, 49))
        style.Font.Weight = ExcelFont.BoldWeight
        style.Font.Color = SpreadsheetColor.FromName(ColorName.White)
        style.WrapText = True
        style.Borders.SetBorders(MultipleBorders.Right Or MultipleBorders.Top, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin)
        worksheet.Cells.GetSubrange("A3:H4").Style = style

        ' Write "Top 10" cells.
        style = New CellStyle()
        style.HorizontalAlignment = HorizontalAlignmentStyle.Center
        style.VerticalAlignment = VerticalAlignmentStyle.Center
        style.Font.Weight = ExcelFont.BoldWeight
        Dim mergedRange = worksheet.Cells.GetSubrange("I5:I14")
        mergedRange.Merged = True
        mergedRange.Value = "T o p   1 0"
        style.Rotation = -90
        style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(198, 239, 206))
        mergedRange.Style = style

        ' Write "Top 20" cells.
        mergedRange = worksheet.Cells.GetSubrange("J5:J24")
        mergedRange.Merged = True
        mergedRange.Value = "T o p   2 0"
        style.IsTextVertical = True
        style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(255, 235, 156))
        mergedRange.Style = style
        mergedRange = worksheet.Cells.GetSubrange("I15:I24")
        mergedRange.Merged = True
        mergedRange.Style = style

        ' Write sample data and formatting to Excel cells.
        For row As Integer = 0 To skyscrapers.GetLength(0) - 1 - 1
            For col As Integer = 0 To skyscrapers.GetLength(1) - 1

                Dim cell = worksheet.Cells(row + 4, col)
                cell.Value = skyscrapers(row + 1, col)

                cell.Style.Borders(IndividualBorder.Right).LineStyle = LineStyle.Thin

                If row Mod 2 = 0 Then cell.Style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(221, 235, 247))

                If col = 0 Then cell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center
                If col > 3 Then cell.Style.Font.Name = "Courier New"
                If col = 4 Then cell.Style.NumberFormat = "#"" m"""
                If col = 5 Then cell.Style.NumberFormat = "#"" ft"""

            Next
        Next

        worksheet.Cells.GetSubrange("A5", "J24").Style.Borders.SetBorders(
            MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Medium)
        worksheet.Cells.GetSubrange("A3", "H4").Style.Borders.SetBorders(
            MultipleBorders.Vertical Or MultipleBorders.Top, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Medium)
        worksheet.Cells.GetSubrange("A5", "I14").Style.Borders.SetBorders(
            MultipleBorders.Bottom Or MultipleBorders.Right, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Medium)

        worksheet.PrintOptions.FitWorksheetWidthToPages = 1

        ' Save workbook as an Excel file.
        workbook.Save("Writing1.xlsx")
    End Sub

    Sub Example2()
        ' Create new empty workbook.
        Dim workbook As New ExcelFile()

        ' Add new sheet.
        Dim worksheet = workbook.Worksheets.Add("Sheet1")
        
        worksheet.Columns("B").SetWidth(400, LengthUnit.Pixel)

        ' Add plain text to cell.
        worksheet.Cells("B2").Value = "This is a plain text."

        ' Add rich formatted text to cell.
        worksheet.Cells("B4").Value = "This is a rich formatted text."
        worksheet.Cells("B4").Style.Font.Color = SpreadsheetColor.FromArgb(255, 128, 64)
        worksheet.Cells("B4").GetCharacters(10, 19).Font.Name = "Arial Black"
        worksheet.Cells("B4").GetCharacters(15, 9).Font.Size = 14 * 20
        worksheet.Cells("B4").GetCharacters(25, 5).Font.Size = 18 * 20

        ' Add HTML formatted text to cells.
        Dim html As String = "<td style='
            font-family: Arial Narrow;
            color: royalblue;
            border: solid black;
            background: #FFF2CC'>This is another rich formatted text.</p>"
        worksheet.Cells("B6").SetValue(html, LoadOptions.HtmlDefault)

        ' Save workbook as an Excel file.
        workbook.Save("Writing2.xlsx")
    End Sub

End Module