using Microsoft.Extensions.Logging;
using Sql2Xls.Interfaces;

namespace Sql2Xls.Excel.Adapters;

public class ExcelExportFactory : IExcelExportFactory
{

    private readonly ILoggerFactory _loggerFactory;
    private readonly ILogger<ExcelExportFactory> _logger;

    public ExcelExportFactory(ILoggerFactory loggerFactory)
    {
        _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
        _logger = _loggerFactory.CreateLogger<ExcelExportFactory>();
    }

    public IExcelExportAdapter CreateAdapter(ExcelExportContext context)
    {
        IExcelExportAdapter excelExport;

        switch (context.ProviderName)
        {
            case "LEGACY":
                excelExport = new ExcelExportAdapter(_loggerFactory.CreateLogger<ExcelExportAdapter>());
                break;

            case "SAX":
                excelExport = new ExcelExportSAXAdapter(_loggerFactory.CreateLogger<ExcelExportSAXAdapter>());
                break;            

            case "ODC":
                excelExport = new ExcelExportODCAdapter(_loggerFactory.CreateLogger<ExcelExportODCAdapter>());
                break;

            case "LARGEXLSX":
                excelExport = new ExcelExportLargeXlsxAdapter(_loggerFactory.CreateLogger<ExcelExportLargeXlsxAdapter>());
                break;

            default:
                excelExport = new ExcelExportSAXAdapter(_loggerFactory.CreateLogger<ExcelExportSAXAdapter>());
                break;
        }

        excelExport.Context = context;

        return excelExport;
    }


}
