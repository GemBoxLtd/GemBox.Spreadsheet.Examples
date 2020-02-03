Imports System.Drawing
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Writing")

        ' Tabular sample data for writing into an Excel file.
        Dim skyscrapers = New Object(20, 7) _
        {
            {"Rank", "Building", "City", "Country", "Metric", "Imperial", "Floors", "Built (Year)"},
            {1, "Burj Khalifa", "Dubai", "United Arab Emirates", 828, 2717, 163, 2010},
            {2, "Shanghai Tower", "Shanghai", "China", 632, 2073, 128, 2015},
            {3, "Abraj Al-Bait Clock Tower", "Mecca", "Saudi Arabia", 601, 1971, 120, 2012},
            {4, "Ping An Finance Centre", "Shenzhen", "China", 599, 1965, 115, 2017},
            {5, "Lotte World Tower", "Seoul", "South Korea", 554.5, 1819, 123, 2016},
            {6, "One World Trade Center", "New York City", "United States", 541.3, 1776, 104, 2014},
            {7, "Guangzhou CTF Finance Centre", "Guangzhou", "China", 530, 1739, 111, 2016},
            {7, "Tianjin CTF Finance Centre", "Tianjin", "China", 530, 1739, 98, 2018},
            {9, "China Zun", "Beijing", "China", 528, 1732, 108, 2018},
            {10, "Taipei 101", "Taipei", "Taiwan", 508, 1667, 101, 2004},
            {11, "Shanghai World Financial Center", "Shanghai", "China", 492, 1614, 101, 2008},
            {12, "International Commerce Centre", "Hong Kong", "China", 484, 1588, 118, 2010},
            {13, "Lakhta Center", "St. Petersburg", "Russia", 462, 1516, 86, 2018},
            {14, "Landmark 81", "Ho Chi Minh City", "Vietnam", 461.2, 1513, 81, 2018},
            {15, "Changsha IFS Tower T1", "Changsha", "China", 452.1, 1483, 88, 2017},
            {16, "Petronas Tower 1", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
            {16, "Petronas Tower 2", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
            {16, "The Exchange 106", "Kuala Lumpur", "Malaysia", 451.9, 1483, 97, 2018},
            {19, "Zifeng Tower", "Nanjing", "China", 450, 1476, 89, 2010},
            {19, "Suzhou IFS", "Suzhou", "China", 450, 1476, 92, 2017}
        }

        worksheet.Cells("A1").Value = "Example of writing typical table - tallest buildings in the world (2019):"

        ' Column width of 8, 30, 16, 20, 9, 11, 9, 9, 4 and 5 characters.
        worksheet.Columns("A").SetWidth(8, LengthUnit.ZeroCharacterWidth) ' Rank
        worksheet.Columns("B").SetWidth(30, LengthUnit.ZeroCharacterWidth) ' Building
        worksheet.Columns("C").SetWidth(16, LengthUnit.ZeroCharacterWidth) ' City
        worksheet.Columns("D").SetWidth(20, LengthUnit.ZeroCharacterWidth) ' Country
        worksheet.Columns("E").SetWidth(9, LengthUnit.ZeroCharacterWidth) ' Metric
        worksheet.Columns("F").SetWidth(11, LengthUnit.ZeroCharacterWidth) ' Imperial
        worksheet.Columns("G").SetWidth(9, LengthUnit.ZeroCharacterWidth) ' Floors
        worksheet.Columns("H").SetWidth(9, LengthUnit.ZeroCharacterWidth) ' Built (Year)
        worksheet.Columns("I").SetWidth(4, LengthUnit.ZeroCharacterWidth)
        worksheet.Columns("J").SetWidth(5, LengthUnit.ZeroCharacterWidth)

        Dim i As Integer, j As Integer
        ' Write header data to Excel cells.
        For j = 0 To 8 - 1 Step j + 1
            worksheet.Cells(3, j).Value = skyscrapers(0, j)
        Next

        worksheet.Cells.GetSubrange("A3:A4").Merged = True ' Rank
        worksheet.Cells.GetSubrange("B3:B4").Merged = True  ' Building
        worksheet.Cells.GetSubrange("C3:C4").Merged = True  ' City
        worksheet.Cells.GetSubrange("D3:D4").Merged = True  ' Country
        worksheet.Cells.GetSubrange("E3:F3").Merged = True ' Height
        worksheet.Cells("E3").Value = "Height"
        worksheet.Cells.GetSubrange("G3:G4").Merged = True  ' Floors
        worksheet.Cells.GetSubrange("H3:H4").Merged = True  ' Built (Year)

        Dim style = New CellStyle
        style.HorizontalAlignment = HorizontalAlignmentStyle.Center
        style.VerticalAlignment = VerticalAlignmentStyle.Center
        style.FillPattern.SetSolid(Color.Chocolate)
        style.Font.Weight = ExcelFont.BoldWeight
        style.Font.Color = Color.White
        style.WrapText = True
        style.Borders.SetBorders(MultipleBorders.Right Or MultipleBorders.Top, Color.Black, LineStyle.Thin)

        worksheet.Cells.GetSubrange("A3:H4").Style = style

        style = New CellStyle
        style.HorizontalAlignment = HorizontalAlignmentStyle.Center
        style.VerticalAlignment = VerticalAlignmentStyle.Center
        style.Font.Weight = ExcelFont.BoldWeight

        Dim mergedRange = worksheet.Cells.GetSubrange("I5:I14")
        mergedRange.Merged = True
        mergedRange.Value = "T o p   1 0"
        style.Rotation = -90
        style.FillPattern.SetSolid(Color.Lime)
        mergedRange.Style = style

        mergedRange = worksheet.Cells.GetSubrange("J5:J24")
        mergedRange.Merged = True
        mergedRange.Value = "T o p   2 0"
        style.IsTextVertical = True
        style.FillPattern.SetSolid(Color.Gold)
        mergedRange.Style = style

        mergedRange = worksheet.Cells.GetSubrange("I15:I24")
        mergedRange.Merged = True
        mergedRange.Style = style

        ' Write and format sample data to Excel cells.
        For i = 0 To 19
            For j = 0 To 7

                Dim cell = worksheet.Cells(i + 4, j)

                cell.Value = skyscrapers(i + 1, j)

                If i Mod 2 = 0 Then
                    cell.Style.FillPattern.SetSolid(Color.LightSkyBlue)
                Else
                    cell.Style.FillPattern.SetSolid(Color.FromArgb(210, 210, 230))
                End If

                If j = 4 Then
                    cell.Style.NumberFormat = "#"" m"""
                End If

                If j = 5 Then
                    cell.Style.NumberFormat = "#"" ft"""
                End If

                If j > 3 Then
                    cell.Style.Font.Name = "Courier New"
                End If

                cell.Style.Borders(IndividualBorder.Right).LineStyle = LineStyle.Thin
            Next j
        Next i

        worksheet.Cells.GetSubrange("A5", "J24").Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Double)
        worksheet.Cells.GetSubrange("A3", "H4").Style.Borders.SetBorders(MultipleBorders.Vertical Or MultipleBorders.Top, Color.Black, LineStyle.Double)
        worksheet.Cells.GetSubrange("A5", "I14").Style.Borders.SetBorders(MultipleBorders.Bottom Or MultipleBorders.Right, Color.Black, LineStyle.Double)

        worksheet.Cells("A27").Value = "Notes:"
        worksheet.Cells("A28").Value = "a) 'Metric' and 'Imperial' columns use custom number formatting."
        worksheet.Cells("A29").Value = "b) All number columns use 'Courier New' font for improved number readability."
        worksheet.Cells("A30").Value = "c) Multiple merged ranges were used for table header and categories header."

        worksheet.PrintOptions.FitWorksheetWidthToPages = 1

        workbook.Save("Writing.xlsx")
    End Sub
End Module