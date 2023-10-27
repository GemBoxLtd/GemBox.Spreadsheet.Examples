using System;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
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

    static void Example2()
    {
        var workbook = ExcelFile.Load("FormControls.xlsx");
        var worksheet = workbook.Worksheets[0];

        // Update CheckBox control.
        var checkBox = worksheet.FormControls[0] as CheckBox;
        checkBox.Checked = false;

        // Read CheckBox control.
        Console.WriteLine("CheckBox checked: " + checkBox.Checked);
        Console.WriteLine("Linked cell value: " + checkBox.CellLink?.Value);
        Console.WriteLine();

        // Update ComboBox control.
        var comboBox = worksheet.FormControls[1] as ComboBox;
        comboBox.SelectedIndex = 1;

        // Read ComboBox control.
        Console.WriteLine("ComboBox range: " + comboBox.InputRange?.Name);
        Console.WriteLine("ComboBox selected: " + comboBox.SelectedValue);
        Console.WriteLine();

        // Update ScrollBar control.
        var scrollBar = worksheet.FormControls[2] as ScrollBar;
        scrollBar.CurrentValue = 33;

        // Read ScrollBar control.
        Console.WriteLine("ScrollBar current: " + scrollBar.CurrentValue);
        Console.WriteLine("Linked cell value: " + scrollBar.CellLink?.Value);
    }
}