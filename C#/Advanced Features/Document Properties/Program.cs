using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        
        var workbook = ExcelFile.Load("ComplexTemplate.xlsx");

        var worksheet = workbook.Worksheets.InsertEmpty(0, "Properties");
        workbook.Worksheets.ActiveWorksheet = worksheet;

        worksheet.Rows[0].Style = workbook.Styles[BuiltInCellStyleName.Heading1];
        worksheet.Columns[0].SetWidth(160, LengthUnit.Pixel);
        worksheet.Columns[1].SetWidth(160, LengthUnit.Pixel);
        worksheet.Columns[2].SetWidth(160, LengthUnit.Pixel);
        worksheet.Columns[3].SetWidth(160, LengthUnit.Pixel);

        worksheet.Cells["A1"].Value = "Built-in Property";
        worksheet.Cells["B1"].Value = "Built-in Value";
        worksheet.Cells["C1"].Value = "Custom Property";
        worksheet.Cells["D1"].Value = "Custom Value";

        int rowIndex = 1;

        // Read built-in document properties.
        foreach (var builtinProperty in workbook.DocumentProperties.BuiltIn)
        {
            worksheet.Cells[rowIndex, 0].Value = builtinProperty.Key.ToString();
            worksheet.Cells[rowIndex, 1].Value = builtinProperty.Value;
            ++rowIndex;
        }

        rowIndex = 1;

        // Read custom document properties.
        foreach (var customProperty in workbook.DocumentProperties.Custom)
        {
            worksheet.Cells[rowIndex, 2].Value = customProperty.Key;
            worksheet.Cells[rowIndex, 3].Value = customProperty.Value;
            ++rowIndex;
        }

        // Write or modify document properties.
        workbook.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Author] = "Jane Doe";
        workbook.DocumentProperties.Custom["Client"] = "New Client";

        workbook.Save("Document Properties.xlsx");
    }
}