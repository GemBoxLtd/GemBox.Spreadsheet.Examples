using System.Collections.Generic;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        
        ExcelFile ef = ExcelFile.Load("TemplateUse.xlsx");

        // Add Sheet
        ExcelWorksheet ws = ef.Worksheets.ActiveWorksheet = ef.Worksheets.InsertEmpty(0, "Document Properties");

        int rowIndex = 0;
        // Read Built-in Document Properies 
        ws.Cells[rowIndex++, 0].Value = "Built-in document properties";

        ws.Cells[rowIndex, 0].Value = "Property";
        ws.Cells[rowIndex++, 1].Value = "Value";

        foreach (KeyValuePair<BuiltInDocumentProperties, string> keyValue in ef.DocumentProperties.BuiltIn)
        {
            ws.Cells[rowIndex, 0].Value = keyValue.Key.ToString();
            ws.Cells[rowIndex++, 1].Value = keyValue.Value;
        }

        // Read Custom Document Properties
        ws.Cells[++rowIndex, 0].Value = "Custom Document Properties";

        ws.Cells[++rowIndex, 0].Value = "Property";
        ws.Cells[rowIndex++, 1].Value = "Value";

        foreach (KeyValuePair<string, object> keyValue in ef.DocumentProperties.Custom)
        {
            ws.Cells[rowIndex, 0].Value = keyValue.Key;
            ws.Cells[rowIndex++, 1].Value = keyValue.Value.ToString();
        }

        // Write/Modifiy Document Properties
        ef.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Author] = "John Doe";
        ef.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Title] = "Genrated title";

        ws.Columns[0].AutoFit();
        ws.Columns[1].AutoFit();

        ef.Save("Document Properties.xlsx");
    }
}
