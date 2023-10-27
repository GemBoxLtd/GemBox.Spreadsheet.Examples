using System;
using System.IO;
using System.Text;
using GemBox.Spreadsheet;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        var workbook = new ExcelFile();
        var worksheet = workbook.Worksheets.Add("Types");

        worksheet.Rows[0].Style = workbook.Styles[BuiltInCellStyleName.Heading1];
        worksheet.Columns[0].Width = 25 * 256;
        worksheet.Columns[1].Width = 25 * 256;
        worksheet.Columns[2].Width = 25 * 256;

        worksheet.Cells[0, 0].Value = "Value";
        worksheet.Cells[0, 1].Value = ".NET Value Type";
        worksheet.Cells[0, 2].Value = "Cell Value Type";

        // Sample data values.
        object[] values =
        {
            DBNull.Value,
            byte.MaxValue,
            sbyte.MinValue,
            short.MinValue,
            ushort.MaxValue,
            1000,
            (uint)2000,
            long.MinValue,
            ulong.MaxValue,
            float.MaxValue,
            double.MaxValue,
            3000.45m,
            true,
            DateTime.Now,
            'a',
            "Sample text.",
            new StringBuilder("Sample text."),
        };

        // Write data and data type to Excel cells.
        for (int i = 0; i < values.Length; i++)
        {
            object value = values[i];

            worksheet.Cells[i + 1, 0].Value = value;
            worksheet.Cells[i + 1, 1].Value = value.GetType().ToString();
        }

        // Save to Excel file and load it back as ExcelFile object.
        using (var stream = new MemoryStream())
        {
            workbook.Save(stream, SaveOptions.XlsxDefault);
            workbook = ExcelFile.Load(stream, LoadOptions.XlsxDefault);
            worksheet = workbook.Worksheets[0];
        }

        // Write cell type to Excel cells.
        for (int i = 0; i < values.Length; i++)
            worksheet.Cells[i + 1, 2].Value = worksheet.Cells[i + 1, 0].ValueType.ToString();

        workbook.Save("Data Types.xlsx");
    }
}
