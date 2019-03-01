using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formula");

        int rowIndex = 0;

        worksheet.Columns[0].Width = 35 * 256;
        worksheet.Columns[1].Width = 15 * 256;
        worksheet.Columns[2].Width = 15 * 256;

        worksheet.Cells[rowIndex++, 0].Value = "Examples of typical formulas usage:";

        worksheet.Cells[++rowIndex, 0].Value = "Some data:";
        worksheet.Cells[rowIndex, 1].SetValue(3);
        worksheet.Cells[rowIndex, 2].SetValue(4.1);
        worksheet.Cells[++rowIndex, 1].SetValue(5.2);
        worksheet.Cells[rowIndex, 2].SetValue(6);
        worksheet.Cells[++rowIndex, 1].SetValue(7);
        worksheet.Cells[rowIndex++, 2].SetValue(8.3);

        // Named ranges.
        string namedRange = "Range1";
        worksheet.NamedRanges.Add(namedRange, worksheet.Cells.GetSubrange("B3", "C4"));

        // Floats without first digit.
        worksheet.Cells[++rowIndex, 0].Value = "Float number without first digit:";
        worksheet.Cells[rowIndex, 1].Formula = "=.5/23+.1-2";

        // Function using named range.
        worksheet.Cells[++rowIndex, 0].Value = "Named range:";
        worksheet.Cells[rowIndex, 1].Formula = "=SUM(" + namedRange + ")";

        // Function's miss argument.
        worksheet.Cells[++rowIndex, 0].Value = "Function's miss arguments:";
        worksheet.Cells[rowIndex, 1].Formula = "=Count(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)";

        // Functions are case-insensitive.
        worksheet.Cells[++rowIndex, 0].Value = "Functions are case-insensitive:";
        worksheet.Cells[rowIndex, 1].Formula = "=cOs( 1 )";

        // Functions.
        worksheet.Cells[++rowIndex, 0].Value = "Supported functions:";

        string nextFunction;
        worksheet.Cells[++rowIndex, 0].Value = "Results";
        worksheet.Cells[rowIndex++, 1].Value = "Formulas";

        nextFunction = "=NOW()+123";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SECOND(12)/23";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MINUTE(24)-1343/35";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=(HOUR(56)-23/35)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=WEEKDAY(5)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=YEAR(23)-WEEKDAY(5)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MONTH(3)-2342/235345";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=((DAY(1)))";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TIME(1,2,3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=DATE(1,2,3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=RAND()";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TEXT(\"text\", \"$d\")";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=VAR(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MOD(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=NOT(FALSE)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=OR(FALSE)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=AND(TRUE)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=FALSE()";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TRUE()";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=VALUE(3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=LEN(\"hello\")";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MID(\"hello\",1,1)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=ROUND(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SIGN(-2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=INT(3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=ABS(-3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=LN(2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=EXP(4)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SQRT(2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=PI()";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=COS(4)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SIN(3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MAX(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MIN(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=AVERAGE(1,2)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SUM(1,3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=IF(1,2,3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=COUNT(1,2,3)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SUBTOTAL(1,B3:C5)";
        worksheet.Cells[rowIndex, 0].Formula = nextFunction;
        worksheet.Cells[rowIndex++, 1].Value = nextFunction;

        // Paranthless checks.
        worksheet.Cells[++rowIndex, 0].Value = "Paranthless:";
        worksheet.Cells[rowIndex, 1].Formula = "=((12+2343+34545))";

        // Unary operators.
        worksheet.Cells[++rowIndex, 0].Value = "Unary operators:";
        worksheet.Cells[rowIndex, 1].Formula = "=B5%";
        worksheet.Cells[rowIndex, 2].Formula = "=+++B5";

        // Operand tokens, bool.
        worksheet.Cells[++rowIndex, 0].Value = "Bool values:";
        worksheet.Cells[rowIndex, 1].Formula = "=TRUE";
        worksheet.Cells[rowIndex, 2].Formula = "=FALSE";

        // Operand tokens, int.
        worksheet.Cells[++rowIndex, 0].Value = "Integer values:";
        worksheet.Cells[rowIndex, 1].Formula = "=1";
        worksheet.Cells[rowIndex, 2].Formula = "=20";

        // Operand tokens, num.
        worksheet.Cells[++rowIndex, 0].Value = "Float values:";
        worksheet.Cells[rowIndex, 1].Formula = "=.4";
        worksheet.Cells[rowIndex, 2].Formula = "=2235.5132";

        // Operand tokens, str.
        worksheet.Cells[++rowIndex, 0].Value = "String values:";
        worksheet.Cells[rowIndex, 1].Formula = "=\"hello world!\"";

        // Operand tokens, error.
        worksheet.Cells[++rowIndex, 0].Value = "Error values:";
        worksheet.Cells[rowIndex, 1].Formula = "=#NULL!";
        worksheet.Cells[rowIndex, 2].Formula = "=#DIV/0!";

        // Binary operators.
        worksheet.Cells[++rowIndex, 0].Value = "Binary operators:";
        worksheet.Cells[rowIndex, 1].Formula = "=(1)-(2)+(3/2+34)/2+12232-32-4";

        workbook.Save("Formula.xlsx");
    }
}
