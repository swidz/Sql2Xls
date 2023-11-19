using Microsoft.Extensions.Logging.Abstractions;
using Sql2Xls.Excel;
using Sql2Xls.Excel.Adapters;
using System.Data;
using System.Diagnostics;

namespace Sql2Xls.Integration.Tests;

//https://github.com/dotnet/Open-XML-SDK/issues/807

[TestClass]
public class Issue_0006_OutOfMemoryException  
{
    private const string _CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

    private string GetRandomString(Random rand, int len)
    {
        return new string(Enumerable.Repeat(_CHARS, len)
            .Select(s => s[rand.Next(s.Length)]).ToArray());
    }

    private DataTable GetDataTable(int numberOfColumns, int numberOfRows, int fieldlen, int seed = 0)
    {
        Random rand = seed == 0 ? new Random() : new Random(seed);

        DataTable dt = new DataTable("MyTable");
        for (int j = 0; j < numberOfColumns; j++)
        {
            dt.Columns.Add($"Column{j}", typeof(string));
        }

        for (int i = 0; i < numberOfRows; i++)
        {
            var row = dt.NewRow();

            for (int j = 0; j < numberOfColumns; j++)
            {
                row[$"Column{j}"] = GetRandomString(rand, fieldlen);
            }

            dt.Rows.Add(row);
        }

        return dt;
    }
    
    [TestMethod]
    [DataRow(154, 250_000, 20, 600)]
    public void T001_GenerateLargeExcel(int numberOfColumns, int numberOfRows, int fieldlen, int seed)
    {
        var dt = GetDataTable(numberOfColumns, numberOfRows, fieldlen, seed);
        
        var start = Stopwatch.StartNew();
        
        using var excelAdapter = new ExcelExportSAXAdapter(NullLogger<ExcelExportSAXAdapter>.Instance)
        {
            Context = new ExcelExportContext()
            {
                FileName = "c:\\datamigration\\excel\\test_001.xlsx",
                SheetName = "MyTable",
                Password = "MyPassword"
            }
        };
        
        excelAdapter.LoadFromDataTable(dt);

        var elapsed = start.Elapsed;
        Console.WriteLine($"Elapsed time: {elapsed.TotalSeconds}");
    }

    [TestMethod]
    [DataRow(154, 250_000, 20, 600)]
    public void T002_GenerateLargeExcel(int numberOfColumns, int numberOfRows, int fieldlen, int seed)
    {
        var dt = GetDataTable(numberOfColumns, numberOfRows, fieldlen, seed);

        var start = Stopwatch.StartNew();
        using var excelAdapter = new ExcelExportSAXAdapterV2(NullLogger<ExcelExportSAXAdapterV2>.Instance)
        {
            Context = new ExcelExportContext()
            {
                FileName = "c:\\datamigration\\excel\\test_002.xlsx",
                SheetName = "MyTable",
                Password = "MyPassword"
            }
        };

        excelAdapter.LoadFromDataTable(dt);

        var elapsed = start.Elapsed;
        Console.WriteLine($"Elapsed time: {elapsed.TotalSeconds}");
    }      
}