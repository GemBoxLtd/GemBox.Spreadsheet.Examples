using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Vba;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        // Create the module.
        VbaModule vbaModule = workbook.VbaProject.Modules.Add(worksheet);
        vbaModule.Code =
@"Sub Button1_Click()
    MsgBox ""Hello World!""
End Sub";

        // Create a button to assign macro.
        var button = worksheet.FormControls.AddButton("Click Me!", "B2", 100, 15, LengthUnit.Point);
        // Assign the macro.
        button.SetMacro(vbaModule, "Button1_Click");

        // Save the workbook as macro-enabled Excel file.
        workbook.Save("AddVbaModule.xlsm");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = ExcelFile.Load("SampleVba.xlsm");

        // Get the module.
        VbaModule vbaModule = workbook.VbaProject.Modules["Module1"];
        // Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Spreadsheet!");

        workbook.Save("UpdateVbaModule.xlsm");
    }
}
