using GemBox.Spreadsheet;
using System;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formulas");

        worksheet.Rows[0].Style = workbook.Styles[BuiltInCellStyleName.Heading1];
        worksheet.Columns[0].Width = 9 * 256;
        worksheet.Columns[1].Width = 36 * 256;
        worksheet.Columns[2].Width = 18 * 256;

        worksheet.Cells[0, 0].Value = "Data";
        worksheet.Cells[0, 1].Value = "Formula";
        worksheet.Cells[0, 2].Value = "Result";

        // Add sample data values.
        worksheet.Cells["A2"].Value = 3;
        worksheet.Cells["A3"].Value = 4.1;
        worksheet.Cells["A4"].Value = 5.2;
        worksheet.Cells["A5"].Value = 6;
        worksheet.Cells["A6"].Value = 7;

        // Add named range.
        worksheet.NamedRanges.Add("MyRange1", worksheet.Cells.GetSubrange("A2:A6"));

        // Sample formulas.
        string[] formulas =
        {
            "=NOW()+123",
            "=MINUTE(0.5)-1343/35",
            "=HOUR(56)-23/35",
            "=YEAR(DATE(2020,1,1)) + 12",
            "=MONTH(3)-2342/235345",
            "=RAND()",
            "=TEXT(\"text\", \"$d\")",
            "=VAR(1,2)",
            "=MOD(1,2)",
            "=NOT(FALSE)",
            "=AND(TRUE)",
            "=TRUE()",
            "=VALUE(3)",
            "=LEN(\"hello\")",
            "=MID(\"hello\",1,1)",
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
            "=SUBTOTAL(1,A2:A4)",                           // Function with cells range.
            "=SUM(MyRange1)",                               // Function with named range.
            "=COUNT(1,  ,  ,,,2, 23,,,,,, 34,,,54,,,,  ,)", // Function with miss argument.
            "=cOs( 1 )",                                    // Functions with different letters case.
            "=+++5",                                        // Unary operators.
            "=(1)-(2)+(3/2+34)/2+12232-32-4",               // Binary operators.
            "=TRUE",                                        // Operand tokens, bool.
            "=20",                                          // Operand tokens, int.
            "=2235.5132",                                   // Operand tokens, num.
            "=\"hello world!\"",                            // Operand tokens, str.
            "=#NULL!"                                       // Operand tokens, error.
        };

        // Write formulas to Excel cells.
        for (int i = 0; i < formulas.Length; i++)
        {
            string formula = formulas[i];
            worksheet.Cells[i + 1, 1].Value = formula;
            worksheet.Cells[i + 1, 2].Formula = formula;
        }

        workbook.Save("Formulas.xlsx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formula Calculation");

        // Some formatting.
        var row = worksheet.Rows[0];
        row.Style.Font.Weight = ExcelFont.BoldWeight;

        var column = worksheet.Columns[0];
        column.SetWidth(250, LengthUnit.Pixel);
        column.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
        column = worksheet.Columns[1];
        column.SetWidth(250, LengthUnit.Pixel);
        column.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

        // Use first row for column headers.
        worksheet.Cells["A1"].Value = "Formula";
        worksheet.Cells["B1"].Value = "Calculated value";

        // Enter some Excel formulas as text in first column.
        worksheet.Cells["A2"].Value = "=1 + 1";
        worksheet.Cells["A3"].Value = "=3 * (2 - 8)";
        worksheet.Cells["A4"].Value = "=3 + ABS(B3)";
        worksheet.Cells["A5"].Value = "=B4 > 15";
        worksheet.Cells["A6"].Value = "=IF(B5, \"Hello world\", \"World hello\")";
        worksheet.Cells["A7"].Value = "=B6 & \" example\"";
        worksheet.Cells["A8"].Value = "=CODE(RIGHT(B7))";
        worksheet.Cells["A9"].Value = "=POWER(B8, 3) * 0.45%";
        worksheet.Cells["A10"].Value = "=SIGN(B9)";
        worksheet.Cells["A11"].Value = "=SUM(B2:B10)";

        // Set text from first column as second row cell's formula.
        int rowIndex = 1;
        while (worksheet.Cells[rowIndex, 0].ValueType != CellValueType.Null)
            worksheet.Cells[rowIndex, 1].Formula = worksheet.Cells[rowIndex++, 0].StringValue;

        // GemBox.Spreadsheet supports single Excel cell calculation, ...
        worksheet.Cells["B2"].Calculate();

        // ... Excel worksheet calculation,
        worksheet.Calculate();

        // ... and whole Excel file calculation.
        worksheet.Parent.Calculate();

        workbook.Save("Formula Calculation.xlsx");
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Formula Evaluation");

        // Enter values in some cells.
        worksheet.Cells["A1"].Value = 1;
        worksheet.Cells["B1"].Value = 2;
        worksheet.Cells["C1"].Value = -1;
        Console.WriteLine($"A1: {worksheet.Cells["A1"].Value}");
        Console.WriteLine($"B1: {worksheet.Cells["B1"].Value}");
        Console.WriteLine($"C1: {worksheet.Cells["C1"].Value}");
        Console.WriteLine();

        // Evaluation of a formula that returns just one value.
        var formula = "=A1 + B1 + C1";
        var value = worksheet.CalculateFormula(formula);

        Console.WriteLine($"Formula: {formula}");
        Console.WriteLine($"Result: {value[0, 0]}");
        Console.WriteLine();

        // Evaluation of a formula that returns more than one value.
        formula = "=ABS(A1:C1)";
        value = worksheet.CalculateFormula(formula);

        Console.WriteLine($"Formula: {formula}");
        for (int i = 0; i < value.GetLength(0); i++)
            for (int j = 0; j < value.GetLength(1); j++)
                Console.WriteLine($"Result [{i}, {j}]: {value[i, j]}");
    }
}
