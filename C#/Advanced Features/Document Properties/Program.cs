using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        
        var workbook = ExcelFile.Load("TemplateUse.xlsx");

        // Add Sheet.
        var worksheet = workbook.Worksheets.ActiveWorksheet = workbook.Worksheets.InsertEmpty(0, "Document Properties");

        int rowIndex = 0;
        // Read Built-in Document Properties.
        worksheet.Cells[rowIndex++, 0].Value = "Built-in document properties";

        worksheet.Cells[rowIndex, 0].Value = "Property";
        worksheet.Cells[rowIndex++, 1].Value = "Value";

        foreach (var keyValue in workbook.DocumentProperties.BuiltIn)
        {
            worksheet.Cells[rowIndex, 0].Value = keyValue.Key.ToString();
            worksheet.Cells[rowIndex++, 1].Value = keyValue.Value;
        }

        // Read Custom Document Properties
        worksheet.Cells[++rowIndex, 0].Value = "Custom Document Properties";

        worksheet.Cells[++rowIndex, 0].Value = "Property";
        worksheet.Cells[rowIndex++, 1].Value = "Value";

        foreach (var keyValue in workbook.DocumentProperties.Custom)
        {
            worksheet.Cells[rowIndex, 0].Value = keyValue.Key;
            worksheet.Cells[rowIndex++, 1].Value = keyValue.Value.ToString();
        }

        // Write/Modify Document Properties.
        workbook.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Author] = "John Doe";
        workbook.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Title] = "Generated title";

        worksheet.Columns[0].SetWidth(192, LengthUnit.Pixel);
        worksheet.Columns[1].SetWidth(217, LengthUnit.Pixel);

        workbook.Save("Document Properties.xlsx");
    }
}