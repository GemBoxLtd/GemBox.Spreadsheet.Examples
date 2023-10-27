using System.IO;
using GemBox.Spreadsheet;

public sealed class LargeCsvReader : TextReader
{
    private const int MaxRow = 1_048_576;
    private readonly TextReader reader;
    private readonly CsvLoadOptions options;

    private int currentRow;
    private bool finished;

    public static ExcelFile ReadFile(string path, CsvLoadOptions options)
    {
        var workbook = new ExcelFile();
        int sheetIndex = 0;

        using (var reader = new LargeCsvReader(path, options))
            while (reader.CanReadNextSheet())
                reader.ReadSheet(workbook, $"Sheet{++sheetIndex}");

        return workbook;
    }

    private LargeCsvReader(string path, CsvLoadOptions options)
    {
        this.reader = File.OpenText(path);
        this.options = options;
    }

    public override string ReadLine()
    {
        if (this.currentRow == MaxRow)
            return null;

        ++this.currentRow;
        string line = this.reader.ReadLine();
        if (line == null)
            this.finished = true;

        return line;
    }

    private void ReadSheet(ExcelFile workbook, string name)
    {
        var worksheet = ExcelFile.Load(this, this.options).Worksheets.ActiveWorksheet;
        workbook.Worksheets.AddCopy(name, worksheet);
    }

    private bool CanReadNextSheet()
    {
        if (this.finished)
            return false;

        this.currentRow = 0;
        return true;
    }

    protected override void Dispose(bool disposing) => this.reader.Dispose();
}