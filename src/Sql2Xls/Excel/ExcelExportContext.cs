using System.ComponentModel;

namespace Sql2Xls.Excel;

public class ExcelExportContext
{
    public static readonly ExcelExportContext _defaultInstance = new ExcelExportContext();
    public static ExcelExportContext Default { get { return _defaultInstance; } }

    public string ProviderName { get; set; }
    public string FileName { get; set; }
    public bool CanIncludeHeader { get; set; }
    public bool DateTimeAsString { get; set; }
    public string SheetName { get; set; }
    public bool CanCreateExtendedFileProperties { get; set; }
    public bool CanCreateCoreFileProperties { get; set; }
    public bool CanUseRelativePaths { get; set; }
    public bool CanRemoveAliasFromDefaultNamespace { get; set; }
    public bool CanCreateThemePart { get; set; }
    public bool CanFixContentTypes { get; set; }
    public bool CanFixXmlDeclarations { get; set; }
    public int MaxRowsPerSheet { get; set; }

    public string ODCConnectionString { get; set; }
    public string ODCSqlStatement { get; set; }
    public string ODCTableName { get; set; }

    public string Password { get; set; }

    public ExcelExportContext()
    {
        ProviderName = "LARGEXLSX";
        CanIncludeHeader = true;        
        SheetName = ExcelConstants.DEFAULT_SHEET_NAME;
        CanCreateExtendedFileProperties = true;
        CanCreateCoreFileProperties = true;
        CanUseRelativePaths = false;
        CanRemoveAliasFromDefaultNamespace = false;
        CanCreateThemePart = false;
        CanFixContentTypes = false;
        CanFixXmlDeclarations = false;
        DateTimeAsString = true;
        MaxRowsPerSheet = ExcelConstants.MAX_ROWS_PER_WORKSHEET;
        ODCTableName = ExcelConstants.DEFAULT_SHEET_NAME;
        ODCConnectionString = string.Empty;
        ODCSqlStatement = string.Empty;
        Password = null;
    }

}
