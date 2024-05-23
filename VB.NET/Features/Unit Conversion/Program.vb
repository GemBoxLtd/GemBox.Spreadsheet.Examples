Imports GemBox.Spreadsheet
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Template.xlsx")
        Dim worksheet = workbook.Worksheets(0)
        Dim cell = worksheet.Cells("A1")

        Dim widthInPoints As Double = cell.Column.GetWidth(LengthUnit.Point)
        Dim heightInPoints As Double = cell.Row.GetHeight(LengthUnit.Point)

        Console.WriteLine("A1 cell's size in different units:")

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            ' The CharacterWidth should not be used with LengthUnitConverter, see:
            ' https://www.gemboxsoftware.com/spreadsheet/docs/GemBox.Spreadsheet.LengthUnit.html
            If unit = LengthUnit.CharacterWidth Then Continue For

            Dim convertedWidth As Double = LengthUnitConverter.Convert(widthInPoints, LengthUnit.Point, unit)
            Dim convertedHeight As Double = LengthUnitConverter.Convert(heightInPoints, LengthUnit.Point, unit)
            Console.WriteLine($"{convertedWidth:0.###} x {convertedHeight:0.###} {unit}")

        Next

    End Sub
End Module
