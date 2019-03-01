Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Dim workbook = New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Formula")

        Dim rowIndex As Integer = 0

        worksheet.Columns(0).Width = 35 * 256
        worksheet.Columns(1).Width = 15 * 256
        worksheet.Columns(2).Width = 15 * 256

        worksheet.Cells(rowIndex, 0).Value = "Examples of typical formulas usage:"
        rowIndex = rowIndex + 2

        worksheet.Cells(rowIndex, 0).Value = "Some data:"
        worksheet.Cells(rowIndex, 1).SetValue(3)
        worksheet.Cells(rowIndex, 2).SetValue(4.1)
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 1).SetValue(5.2)
        worksheet.Cells(rowIndex, 2).SetValue(6)
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 1).SetValue(7)
        worksheet.Cells(rowIndex, 2).SetValue(8.3)
        rowIndex = rowIndex + 1

        ' Named ranges.
        Dim namedRange As String = "Range1"
        worksheet.NamedRanges.Add(namedRange, worksheet.Cells.GetSubrange("B3", "C4"))

        ' Floats without first digit.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Float number without first digit:"
        worksheet.Cells(rowIndex, 1).Formula = "=.5/23+.1-2"

        ' Function using named range.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Named range:"
        worksheet.Cells(rowIndex, 1).Formula = "=SUM(" + namedRange + ")"

        ' Function's miss argument.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Function's miss arguments:"
        worksheet.Cells(rowIndex, 1).Formula = "=Count(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)"

        ' Functions are case-insensitive.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Functions are case-insensitive:"
        worksheet.Cells(rowIndex, 1).Formula = "=cOs( 1 )"

        ' Functions.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Supported functions:"

        Dim nextFunction As String
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Results"
        worksheet.Cells(rowIndex, 1).Value = "Formulas"
        rowIndex = rowIndex + 1

        nextFunction = "=NOW()+123"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SECOND(12)/23"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MINUTE(24)-1343/35"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=(HOUR(56)-23/35)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=WEEKDAY(5)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=YEAR(23)-WEEKDAY(5)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MONTH(3)-2342/235345"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=((DAY(1)))"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=TIME(1,2,3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=DATE(1,2,3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=RAND()"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=TEXT(""text"", ""$d"")"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=VAR(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MOD(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=NOT(FALSE)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=OR(FALSE)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=AND(TRUE)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=FALSE()"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=TRUE()"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=VALUE(3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=LEN(""hello"")"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MID(""hello"",1,1)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=ROUND(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SIGN(-2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=INT(3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=ABS(-3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=LN(2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=EXP(4)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SQRT(2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=PI()"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=COS(4)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SIN(3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MAX(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=MIN(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=AVERAGE(1,2)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SUM(1,3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=IF(1,2,3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=COUNT(1,2,3)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        nextFunction = "=SUBTOTAL(1,B3:C5)"
        worksheet.Cells(rowIndex, 0).Formula = nextFunction
        worksheet.Cells(rowIndex, 1).Value = nextFunction
        rowIndex = rowIndex + 1

        ' Paranthless checks.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Paranthless:"
        worksheet.Cells(rowIndex, 1).Formula = "=((12+2343+34545))"

        ' Unary operators.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Unary operators:"
        worksheet.Cells(rowIndex, 1).Formula = "=B5%"
        worksheet.Cells(rowIndex, 2).Formula = "=+++B5"

        ' Operand tokens, bool.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Bool values:"
        worksheet.Cells(rowIndex, 1).Formula = "=TRUE"
        worksheet.Cells(rowIndex, 2).Formula = "=FALSE"

        ' Operand tokens, int.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Integer values:"
        worksheet.Cells(rowIndex, 1).Formula = "=1"
        worksheet.Cells(rowIndex, 2).Formula = "=20"

        ' Operand tokens, num.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Float values:"
        worksheet.Cells(rowIndex, 1).Formula = "=.4"
        worksheet.Cells(rowIndex, 2).Formula = "=2235.5132"

        ' Operand tokens, str.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "String values:"
        worksheet.Cells(rowIndex, 1).Formula = "=""hello world!"""

        ' Operand tokens, error.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Error values:"
        worksheet.Cells(rowIndex, 1).Formula = "=#NULL!"
        worksheet.Cells(rowIndex, 2).Formula = "=#DIV/0!"

        ' Binary operators.
        rowIndex = rowIndex + 1
        worksheet.Cells(rowIndex, 0).Value = "Binary operators:"
        worksheet.Cells(rowIndex, 1).Formula = "=(1)-(2)+(3/2+34)/2+12232-32-4"

        workbook.Save("Formula.xlsx")
    End Sub
End Module