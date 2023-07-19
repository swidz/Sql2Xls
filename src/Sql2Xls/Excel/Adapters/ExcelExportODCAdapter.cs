using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Sql2Xls.Excel.Parts;
using System.Data;

namespace Sql2Xls.Excel.Adapters;

public class ExcelExportODCAdapter : ExcelExportAdapter
{
    private readonly ILogger<ExcelExportODCAdapter> _logger;
    private ConnectionsPart xlConnectionsPart;

    public ExcelExportODCAdapter(ILogger<ExcelExportODCAdapter> logger) : base(logger)
    {
        _logger = logger;
    }


    public override SpreadsheetDocument Open()
    {
        xlDocument = SpreadsheetDocument.Create(Context.FileName, SpreadsheetDocumentType.Workbook);

        CreateExtendedFileProperties(xlDocument);
        CreateCoreFileProperties(xlDocument);

        xlWorkbookPart = CreateWorkbookPart(xlDocument);

        xlStylesPart = new ExcelStylesPart(xlDocument, workbookStylesPartRelationshipId, Context);
        xlStylesPart.CreateWorkbookStylesPart(xlWorkbookPart);

        xlThemePart = CreateThemePart(xlDocument, xlWorkbookPart);

        //xlSharedStringTablePart = CreateSharedStringTablePart(xlDocument);

        var sheetInfo = CreateSpreadSheetInfo();

        xlWorkbook = CreateWorkbook(xlWorkbookPart, sheetInfo);

        xlWorksheetPart = CreateWorksheetPart(xlWorkbookPart);
        xlWorksheet = CreateWorksheetPre(xlDocument, xlWorksheetPart);

        xlSheetData = new SheetData();
        xlWorksheet.AppendChild(xlSheetData);

        xlConnectionsPart = CreateConnectionsPart(xlWorkbookPart);
        CreateConnection(xlWorksheetPart, xlWorkbookPart, xlConnectionsPart, Context.ODCTableName, Context.ODCConnectionString, Context.ODCSqlStatement);

        CreateWorksheetPost(xlDocument, xlWorksheetPart, xlWorksheet);

        __STATE = STATE_OPEN;

        return xlDocument;
    }

    public override void AddDataRecord(IDataRecord dataRecord)
    {
    }

    public override void Close()
    {
        if (__STATE == STATE_OPEN || __STATE == STATE_IMPORT)
        {
            if (xlDocument != null)
            {
                xlDocument.Dispose();
                xlDocument = null;
            }
            __STATE = STATE_CLOSED;
        }
    }
}
