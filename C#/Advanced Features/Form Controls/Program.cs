using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Form Controls");

        var checkBox = worksheet.FormControls.AddCheckBox("Simple check box", "B2", 100, 15, LengthUnit.Point);
        checkBox.CellLink = worksheet.Cells["A2"];
        checkBox.Checked = true;

        worksheet.Cells["A4"].Value = "VALUE A";
        worksheet.Cells["A5"].Value = "VALUE B";
        worksheet.Cells["A6"].Value = "VALUE C";
        worksheet.Cells["A7"].Value = "VALUE D";
        var comboBox = worksheet.FormControls.AddComboBox("B4", 100, 20, LengthUnit.Point);
        comboBox.InputRange = worksheet.Cells.GetSubrange("A4:A7");
        comboBox.SelectedIndex = 2;

        var scrollBar = worksheet.FormControls.AddScrollBar("B9", 100, 20, LengthUnit.Point);
        scrollBar.CellLink = worksheet.Cells["A9"];
        scrollBar.MinimumValue = 10;
        scrollBar.MaximumValue = 50;
        scrollBar.CurrentValue = 20;

        workbook.Save("Form Controls.xlsx");
    }
}