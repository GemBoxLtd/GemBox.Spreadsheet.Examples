Imports GemBox.Spreadsheet

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Formulas")

        worksheet.Rows(0).Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Columns(0).Width = 9 * 256
        worksheet.Columns(1).Width = 36 * 256
        worksheet.Columns(2).Width = 18 * 256

        worksheet.Cells(0, 0).Value = "Data"
        worksheet.Cells(0, 1).Value = "Formula"
        worksheet.Cells(0, 2).Value = "Result"

        ' Add sample data values.
        worksheet.Cells("A2").Value = 3
        worksheet.Cells("A3").Value = 4.1
        worksheet.Cells("A4").Value = 5.2
        worksheet.Cells("A5").Value = 6
        worksheet.Cells("A6").Value = 7

        ' Add named range.
        worksheet.NamedRanges.Add("MyRange1", worksheet.Cells.GetSubrange("A2:A6"))

        ' Sample formulas.
        Dim formulas As String() =
        {
            "=NOW()+123",
            "=MINUTE(0.5)-1343/35",
            "=HOUR(56)-23/35",
            "=YEAR(DATE(2020,1,1)) + 12",
            "=MONTH(3)-2342/235345",
            "=RAND()",
            "=TEXT(""text"", ""$d"")",
            "=VAR(1,2)",
            "=MOD(1,2)",
            "=NOT(FALSE)",
            "=AND(TRUE)",
            "=TRUE()",
            "=VALUE(3)",
            "=LEN(""hello"")",
            "=MID(""hello"",1,1)",
            "=ROUND(1,2)",
            "=SIGN(-2)",
            "=INT(3)",
            "=ABS(-3)",
            "=LN(2)",
            "=EXP(4)",
            "=SQRT(2)",
            "=PI()",
            "=COS(4)",
            "=MAX(1,2)",
            "=MIN(1,2)",
            "=AVERAGE(1,2)",
            "=SUM(1,3)",
            "=IF(1,2,3)",
            "=COUNT(1,2,3)",
            "=SUBTOTAL(1,A2:A4)",                           ' Function with cells range.
            "=SUM(MyRange1)",                               ' Function with named range.
            "=COUNT(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)", ' Function with miss argument.
            "=cOs( 1 )",                                    ' Functions with different letters case.
            "=+++5",                                        ' Unary operators.
            "=(1)-(2)+(3/2+34)/2+12232-32-4",               ' Binary operators.
            "=TRUE",                                        ' Operand tokens, bool.
            "=20",                                          ' Operand tokens, int.
            "=2235.5132",                                   ' Operand tokens, num.
            "=""hello world!""",                            ' Operand tokens, str.
            "=#NULL!"                                       ' Operand tokens, error.
        }

        ' Write formulas to Excel cells.
        For i = 0 To formulas.Length - 1
            Dim formula As String = formulas(i)
            worksheet.Cells(i + 1, 1).Value = formula
            worksheet.Cells(i + 1, 2).Formula = formula
        Next

        workbook.Save("Formulas.xlsx")
    End Sub

    Sub Example2()
        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("Formulas")

        worksheet.Cells("A1").Value = 4
        worksheet.Cells("A2").Value = 9
        worksheet.Cells("A3").Value = 16
        worksheet.Cells("A4").Value = 25
        worksheet.Cells("A5").Value = 36
        
        ' Set dynamic array formula
        worksheet.Cells("B1").SetDynamicArrayFormula("=SQRT(A1:A5)")
        
        ' Set legacy array formula to C1:C5 range
        worksheet.Cells.GetSubrange("C1:C5").SetArrayFormula("=SQRT(A1:A5)")
            
        ' Set dynamic array formula with a single result
        worksheet.Cells("D1").SetDynamicArrayFormula("=SUM(SQRT(A1:A5))")
            
        ' Set normal formula which will use intersection operator
        worksheet.Cells("E1").Formula = "=SUM(SQRT(A1:A5))"
        
        worksheet.Calculate()
        
        workbook.Save("ArrayFormulas.xlsx")
    End Sub

End Module