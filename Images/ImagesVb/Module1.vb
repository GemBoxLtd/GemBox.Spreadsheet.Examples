Imports System.IO
Imports GemBox.Spreadsheet

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim ef As ExcelFile = New ExcelFile
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("Images")

        Dim pathToResources As String = "Resources"

        ws.Cells(0, 0).Value = "Image examples:"

        ' Small BMP added by using rectangle.
        ws.Pictures.Add(Path.Combine(pathToResources, "SmallImage.bmp"), 50, 50, 48, 48, LengthUnit.Pixel)

        ' Large JPG added by using one anchor.
        ws.Pictures.Add(Path.Combine(pathToResources, "FragonardReader.jpg"), "B9")

        ' PNG added by using two anchors.
        ws.Pictures.Add(Path.Combine(pathToResources, "Dices.png"), "J16", "K20")

        ' GIF added by using anchors. Notice that animation is lost in MS Excel.
        ws.Pictures.Add(Path.Combine(pathToResources, "Zahnrad.gif"),
            New AnchorCell(ws.Columns(9), ws.Rows(21), 100000, 100000),
            New AnchorCell(ws.Columns(10), ws.Rows(23), 50000, 50000)).Position.Mode = PositioningMode.Move

        ' WMF added by using one anchor and size.
        ws.Pictures.Add(Path.Combine(pathToResources, "Graphics1.wmf"), "J9", 250, 100, LengthUnit.Pixel).Position.Mode = PositioningMode.FreeFloating

        ef.Save("Images.xlsx")

    End Sub

End Module