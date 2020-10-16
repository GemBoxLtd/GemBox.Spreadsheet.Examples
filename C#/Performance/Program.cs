using System;
using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using GemBox.Spreadsheet;

[SimpleJob(RuntimeMoniker.Net48)]
[SimpleJob(RuntimeMoniker.NetCoreApp31)]
public class Program
{
    private ExcelFile workbook;
    private readonly Consumer consumer = new Consumer();

    public static void Main()
    {
        BenchmarkRunner.Run<Program>();
    }

    [GlobalSetup]
    public void SetLicense()
    {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

        // If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        // https://www.gemboxsoftware.com/spreadsheet/examples/free-trial-professional-modes/1001

        this.workbook = ExcelFile.Load("RandomSheets.xlsx");
    }

    [Benchmark]
    public ExcelFile Reading()
    {
        return ExcelFile.Load("RandomSheets.xlsx");
    }

    [Benchmark]
    public void Writing()
    {
        using (var stream = new MemoryStream())
            this.workbook.Save(stream, new XlsxSaveOptions());
    }

    [Benchmark]
    public void Iterating()
    {
        this.LoopThroughAllCells().Consume(this.consumer);
    }

    public IEnumerable<object> LoopThroughAllCells()
    {
        foreach (var worksheet in this.workbook.Worksheets)
            foreach (var row in worksheet.Rows)
                foreach (var cell in row.AllocatedCells)
                    yield return cell.Value;
    }
}