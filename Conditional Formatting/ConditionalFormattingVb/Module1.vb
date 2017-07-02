Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.ConditionalFormatting

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Conditional Formatting")

        Dim rowCount As Integer = 20

        ' Specify sheet formatting.
        ws.Rows(0).Style.Font.Weight = ExcelFont.BoldWeight
        ws.Columns(0).SetWidth(3, LengthUnit.Centimeter)
        ws.Columns(1).SetWidth(3, LengthUnit.Centimeter)
        ws.Columns(2).SetWidth(3, LengthUnit.Centimeter)
        ws.Columns(3).SetWidth(3, LengthUnit.Centimeter)
        ws.Columns(3).Style.NumberFormat = "[$$-409]#,##0.00"
        ws.Columns(4).SetWidth(3, LengthUnit.Centimeter)
        ws.Columns(4).Style.NumberFormat = "yyyy-mm-dd"

        Dim cells = ws.Cells

        ' Specify header row.
        cells(0, 0).Value = "Departments"
        cells(0, 1).Value = "Names"
        cells(0, 2).Value = "Years of Service"
        cells(0, 3).Value = "Salaries"
        cells(0, 4).Value = "Deadlines"

        ' Insert random data to sheet.
        Dim random = New Random()
        Dim departments = New String() {"Legal", "Marketing", "Finance", "Planning", "Purchasing"}
        Dim names = New String() {"John Doe", "Fred Nurk", "Hans Meier", "Ivan Horvat"}
        For i As Integer = 0 To rowCount - 1
            cells(i + 1, 0).Value = departments(random.Next(departments.Length))
            cells(i + 1, 1).Value = names(random.Next(names.Length)) + " "c + (i + 1).ToString()
            cells(i + 1, 2).SetValue(random.Next(1, 31))
            cells(i + 1, 3).SetValue(random.Next(10, 101) * 100)
            cells(i + 1, 4).SetValue(DateTime.Now.AddDays(random.Next(-1, 2)))
        Next

        ' Apply shading to alternate rows in a worksheet using 'Formula' based conditional formatting.
        ws.ConditionalFormatting.AddFormula(ws.Cells.Name, "MOD(ROW(),2)=0").
            Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.Accent1Lighter40Pct)
        ws.ConditionalFormatting.AddFormula(ws.Cells.Name, "MOD(ROW(),2)=1").
            Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.Accent5Lighter80Pct)

        ' Apply '2-Color Scale' conditional formatting to 'Years of Service' column.
        ws.ConditionalFormatting.Add2ColorScale("C2:C" & (rowCount + 1))

        ' Apply '3-Color Scale' conditional formatting to 'Salaries' column.
        ws.ConditionalFormatting.Add3ColorScale("D2:D" & (rowCount + 1))

        ' Apply 'Data Bar' conditional formatting to 'Salaries' column.
        ws.ConditionalFormatting.AddDataBar("D2:D" & (rowCount + 1))

        ' Apply 'Icon Set' conditional formatting to 'Years of Service' column.
        ws.ConditionalFormatting.AddIconSet("C2:C" & (rowCount + 1)).IconStyle = SpreadsheetIconStyle.FourTrafficLights

        ' Apply green font color to cells in a 'Years of Service' column which have values between 15 and 20.
        ws.ConditionalFormatting.AddContainValue("C2:C" & (rowCount + 1), ContainValueOperator.Between, 15, 20).
            Style.Font.Color = SpreadsheetColor.FromName(ColorName.Green)

        ' Apply double red border to cells in a 'Names' column which contain text 'Doe'.
        ws.ConditionalFormatting.AddContainText("B2:B" & (rowCount + 1), ContainTextOperator.Contains, "Doe").
            Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Double)

        ' Apply red shading to cells in a 'Deadlines' column which are equal to yesterday's date.
        ws.ConditionalFormatting.AddContainDate("E2:E" & (rowCount + 1), ContainDateOperator.Yesterday).
            Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.Red)

        ' Apply bold font weight to cells in a 'Salaries' column which have top 10 values.
        ws.ConditionalFormatting.AddTopOrBottomRanked("D2:D" & (rowCount + 1), False, 10).
            Style.Font.Weight = ExcelFont.BoldWeight

        ' Apply double underline to cells in a 'Years of Service' column which have below average value.
        ws.ConditionalFormatting.AddAboveOrBelowAverage("C2:C" & (rowCount + 1), True).
            Style.Font.UnderlineStyle = UnderlineStyle.Double

        ' Apply italic font style to cells in a 'Departments' column which have duplicate values.
        ws.ConditionalFormatting.AddUniqueOrDuplicate("A2:A" & (rowCount + 1), True).
            Style.Font.Italic = True

        ef.Save("Conditional Formatting.xlsx")

    End Sub

End Module