Imports System
Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Data Types")

        worksheet.Cells(0, 0).Value = "Cell value examples:"

        ' Column width of 25 and 40 characters.
        worksheet.Columns(0).Width = 25 * 256
        worksheet.Columns(1).Width = 40 * 256

        ' Print gridlines (and show them in PDF, XPS, etc.)
        worksheet.PrintOptions.PrintGridlines = True

        Dim row As Integer = 1

        row = row + 1
        worksheet.Cells(row, 0).Value = "Type"
        worksheet.Cells(row, 1).Value = "Value"

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.DBNull:"
        worksheet.Cells(row, 1).Value = System.DBNull.Value

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Byte:"
        worksheet.Cells(row, 1).SetValue(Byte.MaxValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.SByte:"
        worksheet.Cells(row, 1).SetValue(SByte.MinValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Int16:"
        worksheet.Cells(row, 1).SetValue(Short.MinValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.UInt16:"
        worksheet.Cells(row, 1).SetValue(UShort.MaxValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Int64:"
        worksheet.Cells(row, 1).Value = Long.MinValue

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.UInt64:"
        worksheet.Cells(row, 1).Value = ULong.MaxValue

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.UInt32:"
        worksheet.Cells(row, 1).SetValue(CType(1234, UInteger))

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Int32:"
        worksheet.Cells(row, 1).SetValue(-5678)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Single:"
        worksheet.Cells(row, 1).SetValue(Single.MaxValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Double:"
        worksheet.Cells(row, 1).SetValue(Double.MaxValue)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Boolean:"
        worksheet.Cells(row, 1).SetValue(True)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Char:"
        worksheet.Cells(row, 1).Value = "a"c

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Text.StringBuilder:"
        worksheet.Cells(row, 1).Value = New System.Text.StringBuilder("StringBuilder text.")

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.Decimal:"
        worksheet.Cells(row, 1).Value = 50000D

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.DateTime:"
        worksheet.Cells(row, 1).SetValue(DateTime.Now)

        row = row + 1
        worksheet.Cells(row, 0).Value = "System.String:"
        worksheet.Cells(row, 1).Value = "Microsoft Excel is a spreadsheet program written and distributed by Microsoft for computers using the Microsoft Windows operating system and Apple Macintosh computers. It is overwhelmingly the dominant spreadsheet application available for these platforms and has been so since version 5 1993 and its bundling as part of Microsoft Office." + vbLf _
            + "Microsoft originally marketed a spreadsheet program called Multiplan in 1982, which was very popular on CP/M systems, but on MS-DOS systems it lost popularity to Lotus 1-2-3. This promoted development of a new spreadsheet called Excel which started with the intention to, in the words of Doug Klunder, 'do everything 1-2-3 does and do it better' . The first version of Excel was released for the Mac in 1985 and the first Windows version (numbered 2.0 to line-up with the Mac and bundled with a run-time Windows environment) was released in November 1987. Lotus was slow to bring 1-2-3 to Windows and by 1988 Excel had started to outsell 1-2-3 and helped Microsoft achieve the position of leading PC software developer. This accomplishment, dethroning the king of the software world, solidified Microsoft as a valid competitor and showed its future of developing graphical software. Microsoft pushed its advantage with regular new releases, every two years or so. The current version is Excel 11, also called Microsoft Office Excel 2003." + vbLf _
            + "Early in its life Excel became the target of a trademark lawsuit by another company already selling a software package named 'Excel.' As the result of the dispute Microsoft was required to refer to the program as 'Microsoft Excel' in all of its formal press releases and legal documents. However, over time this practice has slipped." + vbLf _
            + "Excel offers a large number of user interface tweaks, however the essence of UI remains the same as in the original spreadsheet, VisiCalc: the cells are organized in rows and columns, and contain data or formulas with relative or absolute references to other cells." + vbLf _
            + "Excel was the first spreadsheet that allowed the user to define the appearance of spreadsheets (fonts, character attributes and cell appearance). It also introduced intelligent cell recomputation, where only cells dependent on the cell being modified are updated, while previously spreadsheets recomputed everything all the time or waited for a specific user command. Excel has extensive graphing capabilities." + vbLf _
            + "When first bundled into Microsoft Office in 1993, Microsoft Word and Microsoft PowerPoint had their GUIs redesigned for consistency with Excel, the killer app on the PC at the time." + vbLf _
            + "Since 1993 Excel includes support for Visual Basic for Applications (VBA) as a scripting language. VBA is a powerful tool that makes Excel a complete programming environment. VBA and macro recording allow automating routines that otherwise take several manual steps. VBA allows creating forms to handle user input. Automation functionality of VBA exposed Excel as a target for macro viruses." + vbLf _
            + "Excel versions from 5.0 to 9.0 contain various Easter eggs." + vbLf + vbLf + "For more information see: http://en.wikipedia.org/wiki/Microsoft_Excel"

        workbook.Save("Data Types.xlsx")
    End Sub
End Module