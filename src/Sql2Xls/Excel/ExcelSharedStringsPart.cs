using DocumentFormat.OpenXml.Packaging;

namespace Sql2Xls.Excel;


public class ExcelSharedStringsPart : ExcelPart
{
    
    protected readonly Dictionary<string, SharedStringCacheItem> sharedStringsCache;
    public uint Count { get; private set; }
    private uint UniqueCount { get { return (uint)sharedStringsCache.Count; } }

    public ExcelSharedStringsPart(SpreadsheetDocument document,
        string relationshipId,
        ExcelExportContext context,
        WorksheetColumnCollection columns)
        : base(document, relationshipId, context)
    {
        sharedStringsCache = new Dictionary<string, SharedStringCacheItem>(10000);
    }

    public SharedStringTablePart CreateSharedStringTablePart(SpreadsheetDocument document)
    {
        SharedStringTablePart sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(RelationshipId);
        if (Context.CanUseRelativePaths)
        {
            RelationshipId = ExcelHelper.UpdateWorkbookRelationshipsPath(document, sharedStringPart, RelationshipId);
        }
        return sharedStringPart;
    }

    public string FindOrCreate(string valueStr, bool incrementCount = true)
    {
        if (!sharedStringsCache.TryGetValue(valueStr, out SharedStringCacheItem item))
        {
            item = SharedStringCacheItem.Create(sharedStringsCache.Count, valueStr);
            sharedStringsCache.Add(valueStr, item);
        }
        valueStr = item.Position.ToString();

        if (incrementCount)
        {
            Count++;
        }

        return valueStr;
    }
}
