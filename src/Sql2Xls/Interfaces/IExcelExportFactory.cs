using Sql2Xls.Excel;
using Sql2Xls.Excel.Adapters;

namespace Sql2Xls.Interfaces;

public interface IExcelExportFactory
{
    IExcelExportAdapter CreateAdapter(ExcelExportContext context);
}
