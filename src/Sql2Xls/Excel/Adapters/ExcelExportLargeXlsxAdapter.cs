using LargeXlsx;
using Microsoft.Extensions.Logging;
using Sql2Xls.Interfaces;
using System.Data;
using TB.ComponentModel;

namespace Sql2Xls.Excel.Adapters;
public class ExcelExportLargeXlsxAdapter : IExcelExportAdapter, IDisposable
{
    protected WorksheetColumnCollection WorksheetColumns { get; private set; }

    private readonly ILogger<ExcelExportLargeXlsxAdapter> _logger;
    private bool disposedValue;

    private ExcelExportContext _context;
    public ExcelExportContext Context
    {
        get
        {
            if (_context is not null)
                return _context;
            return ExcelExportContext.Default;
        }

        set
        {
            _context = value;
        }
    }

    public ExcelExportLargeXlsxAdapter(ILogger<ExcelExportLargeXlsxAdapter> logger)
    {
        _logger = logger;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {                
            }

            disposedValue = true;
        }
    }

    public void Dispose()
    {     
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    public void AddDataRecord(IDataRecord dataRecord)
    {
        
    }

    public void Close()
    {
        
    }

    public WorksheetColumnCollection InitWorksheetColumns(DataTable dataTable)
    {
        WorksheetColumns = WorksheetColumnCollection.Create(dataTable, Context);
        return WorksheetColumns;
    }

    public void LoadFromDataTable(DataTable dataTable)
    {
        _logger.LogInformation("Generating Microsoft Excel file: {0}", Context.FileName);

        InitWorksheetColumns(dataTable);

        using var stream = new FileStream(Context.FileName, FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream, useZip64: true, requireCellReferences: false);        

        var columns = new XlsxColumn[WorksheetColumns.Count];
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            columns[colIndex] = XlsxColumn.Formatted(30);
            //columns[colIndex] = XlsxColumn.Unformatted();
        }

        xlsxWriter.BeginWorksheet(
            string.IsNullOrEmpty(Context.SheetName) 
                ? Context.SheetName 
                : "Sheet1", 
            columns: columns);
                                            
        _logger.LogTrace("{0} Columns in total", WorksheetColumns.Count);

        xlsxWriter.BeginRow();
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            xlsxWriter.Write(WorksheetColumns[colIndex].ColumnName);
        }

        foreach (DataRow dsrow in dataTable.Rows)
        {            
            xlsxWriter.BeginRow();
            for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
            {
                var val = dsrow[colIndex];
                if (val is null || val == DBNull.Value)
                {                    
                    xlsxWriter.SkipColumns(1);
                    continue;
                }

                var columnInfo = WorksheetColumns[colIndex];
                if (columnInfo.IsSharedString)
                {
                    var stringValue = columnInfo.GetStringValue(val);

                    if (String.IsNullOrEmpty(stringValue))
                    {
                        xlsxWriter.SkipColumns(1);
                        continue;
                    }

                    xlsxWriter.WriteSharedString(stringValue);
                }
                else if(columnInfo.IsDateTime && val.IsConvertibleTo<DateTime>())
                {
                    DateTime dateTimeValue = val.To<DateTime>();
                    xlsxWriter.Write(dateTimeValue);
                }                                
                else if(columnInfo.IsFloat && val.IsConvertibleTo<decimal>())
                {
                    decimal decimalValue = val.To<decimal>();
                    xlsxWriter.Write(decimalValue);
                }
                else if(columnInfo.IsInteger && val.IsConvertibleTo<long>())
                {
                    long longValue = val.To<long>();
                    xlsxWriter.Write((decimal)longValue);
                }
                else if (columnInfo.IsBool && val.IsConvertibleTo<bool>())
                {
                    bool boolValue = val.To<bool>();
                    xlsxWriter.Write(boolValue);
                }
                else
                {
                    var stringValue = columnInfo.GetStringValue(val);
                    xlsxWriter.Write(stringValue);
                }
            }
            
        }        

        _logger.LogTrace("{0} records with {1} columns has been added.", 
            dataTable.Rows.Count, WorksheetColumns.Count);
               
    }    
}
