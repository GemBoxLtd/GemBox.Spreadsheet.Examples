Imports System
Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.Drawing

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Form Controls")

        Dim checkBox = worksheet.FormControls.AddCheckBox("Simple check box", "B2", 100, 15, LengthUnit.Point)
        checkBox.CellLink = worksheet.Cells("A2")
        checkBox.Checked = True

        worksheet.Cells("A4").Value = "VALUE A"
        worksheet.Cells("A5").Value = "VALUE B"
        worksheet.Cells("A6").Value = "VALUE C"
        worksheet.Cells("A7").Value = "VALUE D"
        Dim comboBox = worksheet.FormControls.AddComboBox("B4", 100, 20, LengthUnit.Point)
        comboBox.InputRange = worksheet.Cells.GetSubrange("A4:A7")
        comboBox.SelectedIndex = 2

        Dim scrollBar = worksheet.FormControls.AddScrollBar("B9", 100, 20, LengthUnit.Point)
        scrollBar.CellLink = worksheet.Cells("A9")
        scrollBar.MinimumValue = 10
        scrollBar.MaximumValue = 50
        scrollBar.CurrentValue = 20

        workbook.Save("Form Controls.xlsx")
    End Sub

    Sub Example2()
        Dim workbook = ExcelFile.Load("FormControls.xlsx")
        Dim worksheet = workbook.Worksheets(0)

        ' Update CheckBox control.
        Dim checkBox = TryCast(worksheet.FormControls(0), CheckBox)
        checkBox.Checked = False

        ' Read CheckBox control.
        Console.WriteLine("CheckBox checked: " & checkBox.Checked)
        Console.WriteLine("Linked cell value: " & checkBox.CellLink?.Value)
        Console.WriteLine()

        ' Update ComboBox control.
        Dim comboBox = TryCast(worksheet.FormControls(1), ComboBox)
        comboBox.SelectedIndex = 1

        ' Read ComboBox control.
        Console.WriteLine("ComboBox range: " & comboBox.InputRange?.Name)
        Console.WriteLine("ComboBox selected: " & comboBox.SelectedValue)
        Console.WriteLine()

        ' Update ScrollBar control.
        Dim scrollBar = TryCast(worksheet.FormControls(2), ScrollBar)
        scrollBar.CurrentValue = 33

        ' Read ScrollBar control.
        Console.WriteLine("ScrollBar current: " & scrollBar.CurrentValue)
        Console.WriteLine("Linked cell value: " & scrollBar.CellLink?.Value)
    End Sub

End Module