Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = ExcelFile.Load("Template.xlsx")
        Dim worksheet = workbook.Worksheets(0)
        Dim cell = worksheet.Cells("A1")

        Dim widthInZeroCharacterWidth256thPart As Double = cell.Column.Width
        Dim heightInTwip As Double = cell.Row.Height

        Console.WriteLine("A1 cell's size in different units:")

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            Dim convertedWidth As Double = LengthUnitConverter.Convert(widthInZeroCharacterWidth256thPart, LengthUnit.ZeroCharacterWidth256thPart, unit)
            Dim convertedHeight As Double = LengthUnitConverter.Convert(heightInTwip, LengthUnit.Twip, unit)
            Console.WriteLine($"{convertedWidth:0.###} x {convertedHeight:0.###} {unit}")

        Next

    End Sub
End Module