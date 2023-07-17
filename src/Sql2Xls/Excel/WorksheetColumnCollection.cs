using System.Data;

namespace Sql2Xls.Excel;

public class WorksheetColumnCollection
{
    public int ColumnCount { get; private set; }
    public bool HasSharedStrings { get; private set; }
    public bool DateTimeAsString { get; private set; }
    public WorksheetColumnInfo this[int idx] { get { return innerCollection[idx]; } }

    private readonly List<WorksheetColumnInfo> innerCollection;

    public static WorksheetColumnCollection Create(DataTable dataTable, ExcelExportContext context)
    {
        return new WorksheetColumnCollection(dataTable, context);
    }

    public static WorksheetColumnCollection Create(IDataRecord dataRecord, ExcelExportContext context)
    {
        return new WorksheetColumnCollection(dataRecord, context);
    }

    public static WorksheetColumnCollection Create(ICollection<WorksheetColumnInfo> columns)
    {
        return new WorksheetColumnCollection(columns);
    }

    private WorksheetColumnCollection(DataTable dataTable, ExcelExportContext context)
    {
        ColumnCount = dataTable.Columns.Count;
        innerCollection = new List<WorksheetColumnInfo>(ColumnCount);
        for (int i = 0; i < ColumnCount; i++)
        {
            innerCollection.Add(new WorksheetColumnInfo(dataTable.Columns[i], i, context));
        }
    }

    private WorksheetColumnCollection(IDataRecord dataRecord, ExcelExportContext context)
    {
        ColumnCount = dataRecord.FieldCount;
        innerCollection = new List<WorksheetColumnInfo>(ColumnCount);
        for (int i = 0; i < ColumnCount; i++)
        {
            var columnInfo = new WorksheetColumnInfo(dataRecord, i, context);
            innerCollection.Add(columnInfo);
            if (columnInfo.IsSharedString)
            {
                HasSharedStrings = true;
            }
        }
    }

    private WorksheetColumnCollection(ICollection<WorksheetColumnInfo> columns)
    {
        ColumnCount = columns.Count;
        innerCollection = new List<WorksheetColumnInfo>(ColumnCount);
        foreach (var column in columns)
        {
            innerCollection.Add(column);
            if (column.IsSharedString)
            {
                HasSharedStrings = true;
            }
        }
    }

    public WorksheetColumnInfo GetColumn(int idx)
    {
        return innerCollection[idx];
    }
}
