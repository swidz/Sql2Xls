using Microsoft.Extensions.Logging;

namespace Sql2Xls.Excel;


public interface IExcelExportFactory
{
    ExcelExport Create(ExcelExportContext context);
}

public class ExcelExportFactory : IExcelExportFactory
{

    private readonly ILoggerFactory _loggerFactory;
    private readonly ILogger<ExcelExportFactory> _logger;

    public ExcelExportFactory(ILoggerFactory loggerFactory)
    {
        _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
        _logger = _loggerFactory.CreateLogger<ExcelExportFactory>();
    }

    public ExcelExport Create(ExcelExportContext context)
    {
        ExcelExport excelExport;

        switch (context.ProviderName)
        {
            case "LEGACY":
                excelExport = new ExcelExport(_loggerFactory.CreateLogger<ExcelExport>());
                break;

            case "SAX":
                excelExport = new ExcelExportSAX(_loggerFactory.CreateLogger<ExcelExportSAX>());
                break;

            case "ODC":
                excelExport = new ExcelExportODC(_loggerFactory.CreateLogger<ExcelExportODC>());
                break;

            default:
                excelExport = new ExcelExportSAX(_loggerFactory.CreateLogger<ExcelExportSAX>());
                break;
        }

        excelExport.Context = context;
        
        return excelExport;
    }

    
}
