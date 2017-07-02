using GemBox.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        ExcelFile ef = new ExcelFile();
        ExcelWorksheet ws = ef.Worksheets.Add("Formula");

        int rowIndex = 0;

        ws.Columns[0].Width = 35 * 256;
        ws.Columns[1].Width = 15 * 256;
        ws.Columns[2].Width = 15 * 256;

        ws.Cells[rowIndex++, 0].Value = "Examples of typical formulas usage:";

        ws.Cells[++rowIndex, 0].Value = "Some data:";
        ws.Cells[rowIndex, 1].SetValue(3);
        ws.Cells[rowIndex, 2].SetValue(4.1);
        ws.Cells[++rowIndex, 1].SetValue(5.2);
        ws.Cells[rowIndex, 2].SetValue(6);
        ws.Cells[++rowIndex, 1].SetValue(7);
        ws.Cells[rowIndex++, 2].SetValue(8.3);

        // Named ranges.
        string namedRange = "Range1";
        ws.NamedRanges.Add(namedRange, ws.Cells.GetSubrange("B3", "C4"));

        // Floats without first digit.
        ws.Cells[++rowIndex, 0].Value = "Float number without first digit:";
        ws.Cells[rowIndex, 1].Formula = "=.5/23+.1-2";

        // Function using named range.
        ws.Cells[++rowIndex, 0].Value = "Named range:";
        ws.Cells[rowIndex, 1].Formula = "=SUM(" + namedRange + ")";

        // Function's miss argument.
        ws.Cells[++rowIndex, 0].Value = "Function's miss arguments:";
        ws.Cells[rowIndex, 1].Formula = "=Count(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)";

        // Functions are case-insensitive.
        ws.Cells[++rowIndex, 0].Value = "Functions are case-insensitive:";
        ws.Cells[rowIndex, 1].Formula = "=cOs( 1 )";

        // Functions.
        ws.Cells[++rowIndex, 0].Value = "Supported functions:";

        string nextFunction;
        ws.Cells[++rowIndex, 0].Value = "Results";
        ws.Cells[rowIndex++, 1].Value = "Formulas";

        nextFunction = "=NOW()+123";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SECOND(12)/23";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MINUTE(24)-1343/35";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=(HOUR(56)-23/35)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=WEEKDAY(5)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=YEAR(23)-WEEKDAY(5)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MONTH(3)-2342/235345";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=((DAY(1)))";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TIME(1,2,3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=DATE(1,2,3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=RAND()";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TEXT(\"text\", \"$d\")";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=VAR(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MOD(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=NOT(FALSE)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=OR(FALSE)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=AND(TRUE)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=FALSE()";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=TRUE()";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=VALUE(3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=LEN(\"hello\")";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MID(\"hello\",1,1)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=ROUND(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SIGN(-2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=INT(3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=ABS(-3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=LN(2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=EXP(4)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SQRT(2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=PI()";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=COS(4)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SIN(3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MAX(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=MIN(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=AVERAGE(1,2)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SUM(1,3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=IF(1,2,3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=COUNT(1,2,3)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        nextFunction = "=SUBTOTAL(1,B3:C5)";
        ws.Cells[rowIndex, 0].Formula = nextFunction;
        ws.Cells[rowIndex++, 1].Value = nextFunction;

        // Paranthless checks.
        ws.Cells[++rowIndex, 0].Value = "Paranthless:";
        ws.Cells[rowIndex, 1].Formula = "=((12+2343+34545))";

        // Unary operators.
        ws.Cells[++rowIndex, 0].Value = "Unary operators:";
        ws.Cells[rowIndex, 1].Formula = "=B5%";
        ws.Cells[rowIndex, 2].Formula = "=+++B5";

        // Operand tokens, bool.
        ws.Cells[++rowIndex, 0].Value = "Bool values:";
        ws.Cells[rowIndex, 1].Formula = "=TRUE";
        ws.Cells[rowIndex, 2].Formula = "=FALSE";

        // Operand tokens, int.
        ws.Cells[++rowIndex, 0].Value = "Integer values:";
        ws.Cells[rowIndex, 1].Formula = "=1";
        ws.Cells[rowIndex, 2].Formula = "=20";

        // Operand tokens, num.
        ws.Cells[++rowIndex, 0].Value = "Float values:";
        ws.Cells[rowIndex, 1].Formula = "=.4";
        ws.Cells[rowIndex, 2].Formula = "=2235.5132";

        // Operand tokens, str.
        ws.Cells[++rowIndex, 0].Value = "String values:";
        ws.Cells[rowIndex, 1].Formula = "=\"hello world!\"";

        // Operand tokens, error.
        ws.Cells[++rowIndex, 0].Value = "Error values:";
        ws.Cells[rowIndex, 1].Formula = "=#NULL!";
        ws.Cells[rowIndex, 2].Formula = "=#DIV/0!";

        // Binary operators.
        ws.Cells[++rowIndex, 0].Value = "Binary operators:";
        ws.Cells[rowIndex, 1].Formula = "=(1)-(2)+(3/2+34)/2+12232-32-4";

        ef.Save("Formula.xlsx");
    }
}
