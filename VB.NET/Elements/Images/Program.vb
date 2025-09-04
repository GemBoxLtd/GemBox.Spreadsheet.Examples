Imports GemBox.Spreadsheet
Imports GemBox.Spreadsheet.RichData

Module Program

    Sub Main()

        Example1()
        Example2()
        Example3()

    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

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
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Picture in-cell")

        worksheet.Columns(1).SetWidth(100, LengthUnit.Point)
        worksheet.Rows(1).SetHeight(100, LengthUnit.Point)

        ' Insert an image into a cell.
        worksheet.Cells("B2").RichValue = New RichPictureValue("SmilingFace.png")

        workbook.Save("PictureInCell.xlsx")
    End Sub

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Camera")

        ' Define some data in a specific range of cells.
        worksheet.Cells(0, 0).Value = 100
        worksheet.Cells(0, 1).Value = "ABC"
        worksheet.Cells(1, 0).Value = "DEF"
        worksheet.Cells(1, 1).Value = 200

        ' Add image with camera function enabled.
        worksheet.Pictures.Add("=A1:B2", "E6", "F7")

        workbook.Save("CameraTool.xlsx")
    End Sub

End Module
