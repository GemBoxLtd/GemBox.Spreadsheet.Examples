Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Vba

Module Program

    Sub Main()

        Example1()
        Example2()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Sheet1")

        ' Create the module.
        Dim vbaModule As VbaModule = workbook.VbaProject.Modules.Add(worksheet)
        vbaModule.Code =
"Sub Button1_Click()
    MsgBox ""Hello World!""
End Sub"

        ' Create a button to assign macro.
        Dim button = worksheet.FormControls.AddButton("Click Me!", "B2", 100, 15, LengthUnit.Point)
        ' Assign the macro.
        button.SetMacro(vbaModule, "Button1_Click")

        ' Save the workbook as macro-enabled Excel file.
        workbook.Save("AddVbaModule.xlsm")
    End Sub

    Sub Example2()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("SampleVba.xlsm")

        ' Get the module.
        Dim vbaModule As VbaModule = workbook.VbaProject.Modules("Module1")
        ' Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Spreadsheet!")

        workbook.Save("UpdateVbaModule.xlsm")
    End Sub
End Module
