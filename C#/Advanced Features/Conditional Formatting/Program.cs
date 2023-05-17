using GemBox.Spreadsheet;
using GemBox.Spreadsheet.ConditionalFormatting;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If using the Professional version, remove this FreeLimitReached event handler.
        SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        var workbook = ExcelFile.Load("SampleData.xlsx");
        var worksheet = workbook.Worksheets["Data"];
        int rowCount = worksheet.Rows.Count;

        // Apply shading to alternate rows in a worksheet using 'Formula' based conditional formatting.
        worksheet.ConditionalFormatting.AddFormula(worksheet.Cells.Name, "MOD(ROW(),2)=0")
            .Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.Accent1Lighter40Pct);

        // Apply '2-Color Scale' conditional formatting to 'Years of Service' column.
        worksheet.ConditionalFormatting.Add2ColorScale("C2:C" + (rowCount + 1));

        // Apply '3-Color Scale' conditional formatting to 'Salaries' column.
        worksheet.ConditionalFormatting.Add3ColorScale("D2:D" + (rowCount + 1));

        // Apply 'Data Bar' conditional formatting to 'Salaries' column.
        worksheet.ConditionalFormatting.AddDataBar("D2:D" + (rowCount + 1));

        // Apply 'Icon Set' conditional formatting to 'Years of Service' column.
        worksheet.ConditionalFormatting.AddIconSet("C2:C" + (rowCount + 1)).IconStyle = SpreadsheetIconStyle.FourTrafficLights;

        // Apply green font color to cells in a 'Years of Service' column which have values between 15 and 20.
        worksheet.ConditionalFormatting.AddContainValue("C2:C" + (rowCount + 1), ContainValueOperator.Between, 15, 20)
            .Style.Font.Color = SpreadsheetColor.FromName(ColorName.Green);

        // Apply double red border to cells in a 'Names' column which contain text 'Doe'.
        worksheet.ConditionalFormatting.AddContainText("B2:B" + (rowCount + 1), ContainTextOperator.Contains, "Doe")
            .Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Double);

        // Apply red shading to cells in a 'Deadlines' column which are equal to yesterday's date.
        worksheet.ConditionalFormatting.AddContainDate("E2:E" + (rowCount + 1), ContainDateOperator.Yesterday)
            .Style.FillPattern.PatternBackgroundColor = SpreadsheetColor.FromName(ColorName.Red);

        // Apply bold font weight to cells in a 'Salaries' column which have top 10 values.
        worksheet.ConditionalFormatting.AddTopOrBottomRanked("D2:D" + (rowCount + 1), false, 10)
            .Style.Font.Weight = ExcelFont.BoldWeight;

        // Apply double underline to cells in a 'Years of Service' column which have below average value.
        worksheet.ConditionalFormatting.AddAboveOrBelowAverage("C2:C" + (rowCount + 1), true)
            .Style.Font.UnderlineStyle = UnderlineStyle.Double;

        // Apply italic font style to cells in a 'Departments' column which have duplicate values.
        worksheet.ConditionalFormatting.AddUniqueOrDuplicate("A2:A" + (rowCount + 1), true)
            .Style.Font.Italic = true;

        workbook.Save("Conditional Formatting.xlsx");
    }
}