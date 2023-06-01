Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Images")

        ' Add small BMP image with specified rectangle position.
        worksheet.Pictures.Add("SmallImage.bmp", 50, 50, 48, 48, LengthUnit.Pixel)

        ' Add large JPG image with specified top-left cell.
        worksheet.Pictures.Add("FragonardReader.jpg", "B9")

        ' Add PNG image with specified top-left and bottom-right cells.
        worksheet.Pictures.Add("Dices.png", "J16", "K20")

        ' Add GIF image using anchors.
        Dim picture = worksheet.Pictures.Add("Zahnrad.gif",
            New AnchorCell(worksheet.Columns(9), worksheet.Rows(21), 100000, 100000),
            New AnchorCell(worksheet.Columns(10), worksheet.Rows(23), 50000, 50000))

        ' Set picture's position mode.
        picture.Position.Mode = PositioningMode.Move

        ' Add SVG image with specified top-left cell and size.
        picture = worksheet.Pictures.Add("Graphics1.svg", "J9", 250, 100, LengthUnit.Pixel)

        ' Set picture's metadata.
        picture.Metadata.Name = "SVG Image"

        workbook.Save("Images.xlsx")
    End Sub

    Sub Example2()
        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Smileys")

        ' Create a sheet with specified columns width and rows height.
        For i As Integer = 0 To 5
            worksheet.Columns(i).SetWidth(10 * (i + 1), LengthUnit.Point)
            worksheet.Rows(i).SetHeight(10 * (i + 1), LengthUnit.Point)
        Next

        ' Add images that fit inside a single cell.
        For Each cell In worksheet.Cells.GetSubrange("A1:F6")
            Dim picture = worksheet.Pictures.Add("SmilingFace.png", cell.Name)
            Dim position = picture.Position

            Dim maxWidth As Double = cell.Column.GetWidth(LengthUnit.Point)
            Dim maxHeight As Double = cell.Row.GetHeight(LengthUnit.Point)

            Dim ratioX = maxWidth / position.Width
            Dim ratioY = maxHeight / position.Height
            Dim ratio = Math.Min(ratioX, ratioY)

            If ratio < 1 Then
                position.Width *= ratio
                position.Height *= ratio
            End If
        Next

        workbook.Save("CellsImages.xlsx")
    End Sub

End Module