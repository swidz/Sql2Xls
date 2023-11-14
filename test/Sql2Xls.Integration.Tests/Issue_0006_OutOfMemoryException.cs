using Microsoft.Extensions.Logging.Abstractions;
using Sql2Xls.Excel;
using Sql2Xls.Excel.Adapters;
using System.Data;

namespace Sql2Xls.Integration.Tests;

[TestClass]
public class Issue_0006_OutOfMemoryException  
{
    private const string _CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

    private string GetRandomString(Random rand, int len)
    {
        return new string(Enumerable.Repeat(_CHARS, len)
            .Select(s => s[rand.Next(s.Length)]).ToArray());
    }
    
    [TestMethod]
    [DataRow(154, 250_000, 20)]
    public void GenerateLargeExcel(int numberOfColumns, int numberOfRows, int fieldlen)
    {
        Random rand = new Random();
        
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

        var context = new ExcelExportContext()
        {
            FileName = "c:\\datamigration\\excel\\test.xlsx",
            SheetName = "MyTable"
        };

        using var excelAdapter = new ExcelExportSAXAdapter(NullLogger<ExcelExportSAXAdapter>.Instance);
        excelAdapter.Context = context;
        excelAdapter.LoadFromDataTable(dt);
    }
}