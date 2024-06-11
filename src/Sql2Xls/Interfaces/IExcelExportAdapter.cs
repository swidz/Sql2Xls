using Sql2Xls.Excel;
using System.Data;

namespace Sql2Xls.Interfaces;

public interface IExcelExportAdapter
{
    ExcelExportContext Context { get; set; }

    void AddDataRecord(IDataRecord dataRecord);
    void Close();

    void LoadFromDataTable(DataTable dataTable);
}
