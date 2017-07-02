using System;
using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Data Validation");

        ws.Cells[0, 0].Value = "Data validation examples:";

        ws.Cells[2, 1].Value = "Decimal greater than 3.14 (on entire row 4):";
        ws.DataValidations.Add(new DataValidation(ws.Rows[3].Cells)
        {
            Type = DataValidationType.Decimal,
            Operator = DataValidationOperator.GreaterThan,
            Formula1 = 3.14,
            InputMessageTitle = "Enter a decimal",
            InputMessage = "Decimal should be greater than 3.14.",
            ErrorTitle = "Invalid decimal",
            ErrorMessage = "Value should be a decimal greater than 3.14."
        });
        ws.Cells.GetSubrange("A4", "J4").Value = 3.15;

        ws.Cells[7, 1].Value = "List from B9 to B12 (on cell C8):";
        ws.Cells[8, 1].Value = "John";
        ws.Cells[9, 1].Value = "Fred";
        ws.Cells[10, 1].Value = "Hans";
        ws.Cells[11, 1].Value = "Ivan";
        ws.DataValidations.Add(new DataValidation(ws, "C8")
        {
            Type = DataValidationType.List,
            Formula1 = "=B9:B12",
            InputMessageTitle = "Enter a name",
            InputMessage = "Name should be from the list: John, Fred, Hans, Ivan.",
            ErrorStyle = DataValidationErrorStyle.Warning,
            ErrorTitle = "Invalid name",
            ErrorMessage = "Value should be a name from the list: John, Fred, Hans, Ivan."
        });
        ws.Cells["C8"].Value = "John";

        ws.Cells[13, 1].Value = "Date between 2011-01-01 and 2011-12-31 (on cell range C14:E15):";
        ws.DataValidations.Add(new DataValidation(ws.Cells.GetSubrange("C14", "E15"))
        {
            Type = DataValidationType.Date,
            Operator = DataValidationOperator.Between,
            Formula1 = new DateTime(2011, 1, 1),
            Formula2 = new DateTime(2011, 12, 31),
            InputMessageTitle = "Enter a date",
            InputMessage = "Date should be between 2011-01-01 and 2011-12-31.",
            ErrorStyle = DataValidationErrorStyle.Information,
            ErrorTitle = "Invalid date",
            ErrorMessage = "Value should be a date between 2011-01-01 and 2011-12-31."
        });
        ws.Cells.GetSubrange("C14", "E15").Value = new DateTime(2011, 1, 1);

        // Column width of 8, 55 and 15 characters.
        ws.Columns[0].Width = 8 * 256;
        ws.Columns[1].Width = 55 * 256;
        ws.Columns[2].Width = 15 * 256;
        ws.Columns[3].Width = 15 * 256;
        ws.Columns[4].Width = 15 * 256;

        ef.Save("Data Validation.xlsx");
    }
}
