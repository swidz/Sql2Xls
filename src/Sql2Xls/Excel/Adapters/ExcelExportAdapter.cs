using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Sql2Xls.Excel.Extensions;
using Sql2Xls.Excel.Parts;
using Sql2Xls.Extensions;
using Sql2Xls.Interfaces;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Sql2Xls.Excel.Adapters;

public class ExcelExportAdapter : IExcelExportAdapter, IDisposable
{
    bool disposed = false;

    protected byte __STATE = 0;
    protected const byte STATE_NONE = 0;
    protected const byte STATE_OPEN = 1;
    protected const byte STATE_IMPORT = 2;
    protected const byte STATE_CLOSED = 3;

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

    protected WorksheetColumnCollection WorksheetColumns { get; private set; }

    protected SpreadsheetDocument xlDocument;
    protected SharedStringTablePart xlSharedStringTablePart;
    protected WorkbookPart xlWorkbookPart;
    protected int xlSharedStringsCount;
    protected ExcelStylesPart xlStylesPart;
    protected ExcelThemePart xlThemePart;
    protected WorksheetPart xlWorksheetPart;
    protected Worksheet xlWorksheet;
    protected Workbook xlWorkbook;
    protected SheetData xlSheetData;

    protected int currentRow = 0;

    protected readonly Dictionary<string, SharedStringCacheItem> sharedStringsCache = new(10000);

    protected string workbookPartRelationshipId = "rId1";
    protected string coreFilePropertiesPartRelationshipId = "rId2";
    protected string extendedFilePropertiesPartRelationshipId = "rId3";
    protected string worksheetPartRelationshipId = "rId1";
    protected string themePartRelationshipId = "rId2";
    protected string workbookStylesPartRelationshipId = "rId3";
    protected string sharedStringPartRelationshipId = "rId4";

    private readonly ILogger<ExcelExportAdapter> _logger;

    public ExcelExportAdapter(ILogger<ExcelExportAdapter> logger)
    {
        _logger = logger;
    }


    public WorksheetColumnCollection InitWorksheetColumns(DataTable dataTable)
    {
        WorksheetColumns = WorksheetColumnCollection.Create(dataTable, Context);
        return WorksheetColumns;
    }

    public WorksheetColumnCollection InitWorksheetColumns(IDataRecord record)
    {
        WorksheetColumns = WorksheetColumnCollection.Create(record, Context);
        return WorksheetColumns;
    }

    public void LoadFromDataTable(DataTable dataTable)
    {
        _logger.LogInformation("Generating Microsoft Excel file: {0}", Context.FileName);

        InitWorksheetColumns(dataTable);

        using (SpreadsheetDocument document = SpreadsheetDocument.Create(
            Context.FileName, SpreadsheetDocumentType.Workbook))
        {
            CreateFromDataTable(document, dataTable);
        }

        UpdateExcelArchive(Context.FileName);
    }

    public virtual SpreadsheetDocument Open()
    {
        xlDocument = SpreadsheetDocument.Create(Context.FileName, SpreadsheetDocumentType.Workbook);

        CreateExtendedFileProperties(xlDocument);
        CreateCoreFileProperties(xlDocument);

        xlWorkbookPart = CreateWorkbookPart(xlDocument);

        xlStylesPart = new ExcelStylesPart(xlDocument, workbookStylesPartRelationshipId, Context);
        xlStylesPart.CreateWorkbookStylesPart(xlWorkbookPart);

        xlThemePart = CreateThemePart(xlDocument, xlWorkbookPart);

        xlSharedStringTablePart = CreateSharedStringTablePart(xlDocument);

        var sheetInfo = CreateSpreadSheetInfo();

        xlWorkbook = CreateWorkbook(xlWorkbookPart, sheetInfo);

        xlWorksheetPart = CreateWorksheetPart(xlWorkbookPart);
        xlWorksheet = CreateWorksheetPre(xlDocument, xlWorksheetPart);
        xlSheetData = new SheetData();

        __STATE = STATE_OPEN;

        return xlDocument;
    }

    protected virtual void AddHeaderRow(int rowIndex = 0)
    {
        Row newRow = CreateHeaderRow(rowIndex, true);
        xlSheetData.AppendChild(newRow);
        __STATE = STATE_IMPORT;
    }

    public virtual void AddDataRecord(IDataRecord dataRecord)
    {
        if (__STATE == STATE_NONE)
        {
            InitWorksheetColumns(dataRecord);
            Open();

            if (Context.CanIncludeHeader)
            {
                AddHeaderRow(currentRow++);
            }
        }

        Row newRow = CreateRowFromRecord(dataRecord, currentRow++, true);
        xlSheetData.AppendChild(newRow);
    }

    public virtual void Close()
    {
        if (__STATE == STATE_OPEN || __STATE == STATE_IMPORT)
        {
            if (xlDocument != null)
            {
                CreateSharedStringTable(xlDocument, xlSharedStringTablePart, sharedStringsCache, xlSharedStringsCount);

                if (xlSheetData != null && xlWorksheet != null)
                {
                    xlWorksheet.AppendChild(xlSheetData);
                    CreateWorksheetPost(xlDocument, xlWorksheetPart, xlWorksheet);
                }

                xlDocument.Dispose();
                xlDocument = null;
            }
            __STATE = STATE_CLOSED;
        }
    }

    private void CreateFromDataTable(SpreadsheetDocument document, DataTable dataTable)
    {
        CreateExtendedFileProperties(document);
        CreateCoreFileProperties(document);

        var workbookPart = CreateWorkbookPart(document);
        var stylesPart = CreateWorkbookStylesPart(document, workbookPart);
        var themePart = CreateThemePart(document, workbookPart);
        var sharedStringTablePart = CreateSharedStringTablePart(document);
        CreateSharedStringTable(document, sharedStringTablePart, dataTable);
        var sheetInfo = CreateSpreadSheetInfo();
        var workBook = CreateWorkbook(workbookPart, sheetInfo);
        var worksheetPart = CreateWorksheetPart(workbookPart);

        //TODO Make ExcelExport class hierarchy stateless
        /* Set local variables */
        xlDocument = document;
        //xlSharedStringTablePart;
        xlWorkbookPart = workbookPart;
        //xlSharedStringsCount;
        xlStylesPart = stylesPart;
        xlThemePart = themePart;
        xlWorksheetPart = worksheetPart;
        //xlWorksheet = ;
        xlWorkbook = workBook;
        //xlSheetData;


        CreateWorksheet(document, worksheetPart, dataTable);


    }

    public void LoadFromList<T>(List<T> list)
    {
        LoadFromDataTable(list.ToDataTable());
    }

    protected List<Sheet> CreateSpreadSheetInfo()
    {
        return new List<Sheet>
    {
        new Sheet()
        {
            Name = StringValue.FromString(Context.SheetName),
            SheetId = UInt32Value.FromUInt32(1U),
            Id = workbookPartRelationshipId
        }
    };
    }

    protected string GetValue(object value, WorksheetColumnInfo columnInfo)
    {
        string strValue = value.ToString();
        string resultValue = strValue;

        if (columnInfo.IsFloat)
        {
            if (double.TryParse(strValue, out double doubleValue))
            {
                resultValue = doubleValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        else if (columnInfo.IsDate)
        {
            if (DateTime.TryParse(strValue, out DateTime dateValue))
            {
                if (Context.DateTimeAsString)
                {
                    resultValue = dateValue.ToString(ApplicationConstants.DateTimeFormatString);
                }
                else
                {
                    //xls compliant
                    //double oaValue = dateValue.ToOADate();
                    //resultValue = oaValue.ToString(CultureInfo.InvariantCulture);

                    //xlsx transitional compliant
                    resultValue = dateValue.ToString("s");
                }
            }
        }

        return resultValue;
    }

    protected virtual void CreateExtendedFileProperties(SpreadsheetDocument document)
    {
        if (!Context.CanCreateExtendedFileProperties)
            return;

        new ExcelExtendedFilePropertiesPart(document, extendedFilePropertiesPartRelationshipId, Context).CreateDOM();
    }

    protected void CreateCoreFileProperties(SpreadsheetDocument document)
    {
        if (!Context.CanCreateCoreFileProperties)
            return;

        new ExcelCoreFilePropertiesPart(document, coreFilePropertiesPartRelationshipId, Context).CreateDOM();
    }

    protected WorkbookPart CreateWorkbookPart(SpreadsheetDocument document)
    {
        WorkbookPart workbookPart = document.AddWorkbookPart();
        document.ChangeIdOfPart(workbookPart, workbookPartRelationshipId);
        if (Context.CanUseRelativePaths)
        {
            workbookPartRelationshipId = document.UpdateDocumentRelationshipsPath(workbookPart, ExcelConstants.OfficeDocumentRelationshipType);
        }

        return workbookPart;
    }

    protected WorksheetPart CreateWorksheetPart(WorkbookPart workbookPart)
    {
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(worksheetPartRelationshipId);
        return worksheetPart;
    }

    protected virtual Workbook CreateWorkbook(WorkbookPart workbookPart, List<Sheet> sheetsInfo)
    {
        Workbook workbook = new Workbook();
        workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        FileVersion fileVersion1 = new FileVersion
        {
            ApplicationName = "xl",
            LastEdited = "6",
            LowestEdited = "5",
            BuildVersion = "14420"
            //CodeName = "{7A2D7E96-6E34-419A-AE5F-296B3A7E7977}" 
        };

        WorkbookProperties workbookProperties = new WorkbookProperties
        {
            CodeName = "ThisWorkbook",
            DefaultThemeVersion = UInt32Value.FromUInt32(124226U)
        };

        var bookViews = new BookViews();
        bookViews.Append(new WorkbookView());

        workbook.Append(fileVersion1);
        workbook.Append(workbookProperties);
        workbook.Append(bookViews);

        Sheets sheets = new Sheets();
        foreach (var sheet in sheetsInfo)
        {
            sheets.AppendChild(sheet);
        }

        workbook.Append(sheets);

        CalculationProperties calculationProperties = new CalculationProperties
        {
            FullCalculationOnLoad = new BooleanValue(true)
        };
        workbook.Append(calculationProperties);

        workbookPart.Workbook = workbook;
        return workbook;
    }

    protected virtual Worksheet CreateWorksheetPre(SpreadsheetDocument document, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = new Worksheet
        {
            MCAttributes = new MarkupCompatibilityAttributes()
            {
                Ignorable = "x14ac xr xr2 xr3"
            }
        };

        worksheet.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0000-000000000000}"));
        worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        worksheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
        worksheet.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
        worksheet.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");

        if (Context.CanUseRelativePaths)
        {
            worksheetPartRelationshipId = document.UpdateWorkbookRelationshipsPath(worksheetPart, ExcelConstants.WorksheetRelationshipType);
        }

        /*
        SheetDimension sheetDimension = new SheetDimension
        {
            Reference = StringValue.FromString($"A1:{GetCellReference(WorksheetColumns.ColumnCount - 1, dataTable.Rows.Count + 1)}")
        };
        */

        SheetViews sheetViews = new SheetViews();
        SheetView sheetView = new SheetView
        {
            TabSelected = BooleanValue.FromBoolean(true),
            WorkbookViewId = UInt32Value.FromUInt32(0U)
        };
        sheetViews.Append(sheetView);

        SheetFormatProperties sheetFormatProperties = new SheetFormatProperties
        {
            DefaultRowHeight = DoubleValue.FromDouble(15D),
            DyDescent = DoubleValue.FromDouble(0.25D)
        };

        Columns xlColumns = new Columns();
        Column xlColumn = new Column
        {
            Min = UInt32Value.FromUInt32(1U),
            Max = UInt32Value.FromUInt32((uint)WorksheetColumns.Count),
            Width = DoubleValue.FromDouble(20D),
            CustomWidth = BooleanValue.FromBoolean(true)
        };
        xlColumns.Append(xlColumn);

        //worksheet.Append(sheetDimension);
        worksheet.Append(sheetViews);
        worksheet.Append(sheetFormatProperties);
        worksheet.Append(xlColumns);

        return worksheet;
    }

    protected virtual void CreateWorksheetPost(SpreadsheetDocument document, WorksheetPart worksheetPart, Worksheet worksheet)
    {
        PageMargins pageMargins = new PageMargins
        {
            Left = DoubleValue.FromDouble(0.7D),
            Right = DoubleValue.FromDouble(0.7D),
            Top = DoubleValue.FromDouble(0.75D),
            Bottom = DoubleValue.FromDouble(0.75D),
            Header = DoubleValue.FromDouble(0.3D),
            Footer = DoubleValue.FromDouble(0.3D)
        };

        worksheet.Append(pageMargins);

        HeaderFooter headerFooter = new HeaderFooter();
        worksheet.Append(headerFooter);

        worksheetPart.Worksheet = worksheet;
        worksheetPart.Worksheet.Save();
    }

    protected virtual void CreateWorksheet(SpreadsheetDocument document, WorksheetPart worksheetPart, DataTable dataTable)
    {
        Worksheet worksheet = CreateWorksheetPre(document, worksheetPart);
        CreateSheetData(document, worksheet, dataTable);
        CreateWorksheetPost(document, worksheetPart, worksheet);
    }

    protected ExcelThemePart CreateThemePart(SpreadsheetDocument document, WorkbookPart workbookPart)
    {
        if (!Context.CanCreateThemePart)
            return null;

        ExcelThemePart themePart = new ExcelThemePart(document, themePartRelationshipId, Context);
        themePart.CreateThemePart(workbookPart);
        return themePart;
    }

    protected ExcelStylesPart CreateWorkbookStylesPart(SpreadsheetDocument document, WorkbookPart workbookPart)
    {
        ExcelStylesPart stylesPart = new ExcelStylesPart(document, workbookStylesPartRelationshipId, Context);
        stylesPart.CreateWorkbookStylesPart(workbookPart);
        return stylesPart;
    }

    protected void FixContentTypes(ZipArchive archive)
    {
        _logger.LogTrace("Updating {0}", ExcelConstants.ContentTypesFilename);

        var entry = archive.GetEntry(ExcelConstants.ContentTypesFilename);

        //Replace the content
        StringBuilder sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />");
        sb.Append("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />");
        if (Context.CanCreateThemePart)
            sb.Append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\" />");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />");
        sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" />");
        if (Context.CanCreateCoreFileProperties)
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />");
        if (Context.CanCreateExtendedFileProperties)
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" />");
        sb.Append("</Types>");

        entry.Delete();
        entry = archive.CreateEntry(ExcelConstants.ContentTypesFilename);
        using (StreamWriter writer = new StreamWriter(entry.Open()))
        {
            writer.Write(sb);
        }
    }

    protected void RemoveAliasFromXMLAttributes(ZipArchive archive, string entryName)
    {
        _logger.LogTrace("Updateing {0}", entryName);
        ZipArchiveEntry entry = archive.GetEntry(entryName);
        StringBuilder sb;
        using (StreamReader reader = new StreamReader(entry.Open()))
        {
            sb = new StringBuilder(reader.ReadToEnd());
        }
        entry.Delete();
        entry = archive.CreateEntry(entryName);
        sb = new StringBuilder(GetXMLWithDefaultNamespace(sb.ToString()));
        using (StreamWriter writer = new StreamWriter(entry.Open()))
        {
            writer.Write(sb);
        }
    }

    protected void StandaloneXmlDeclarations(ZipArchive archive, string entryName)
    {
        _logger.LogTrace("Updateing {0}", entryName);

        ZipArchiveEntry entry = archive.GetEntry(entryName);
        StringBuilder sb;
        using (StreamReader reader = new StreamReader(entry.Open()))
        {
            sb = new StringBuilder(reader.ReadToEnd());
        }
        entry.Delete();
        entry = archive.CreateEntry(entryName);
        var xml = XDocument.Parse(sb.ToString());
        xml.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
        sb = new StringBuilder(xml.ToStringWithDeclaration());
        using (StreamWriter writer = new StreamWriter(entry.Open()))
        {
            writer.Write(sb);
        }
    }

    private void UpdateExcelArchive(string filePath)
    {
        if (Context.CanFixContentTypes || Context.CanFixXmlDeclarations || Context.CanRemoveAliasFromDefaultNamespace)
        {
            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Update))
            {
                if (Context.CanFixContentTypes)
                {
                    FixContentTypes(archive);
                }

                if (Context.CanRemoveAliasFromDefaultNamespace)
                {
                    RemoveAliasFromXMLAttributes(archive, "xl/sharedStrings.xml");
                    RemoveAliasFromXMLAttributes(archive, "xl/styles.xml");
                    RemoveAliasFromXMLAttributes(archive, "xl/workbook.xml");
                    RemoveAliasFromXMLAttributes(archive, "xl/worksheets/sheet1.xml");
                }

                if (Context.CanFixXmlDeclarations)
                {
                    StandaloneXmlDeclarations(archive, "xl/_rels/workbook.xml.rels");
                    StandaloneXmlDeclarations(archive, "_rels/.rels");
                }
            }
        }
    }

    protected virtual Row CreateHeaderRow(int rowIndex = 0, bool preCacheSharedString = true)
    {
        Cell[] cellChildren = new Cell[WorksheetColumns.Count];
        var headerRow = new Row { RowIndex = (uint)rowIndex + 1 };
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            string valueStr = columnInfo.ColumnName;

            if (preCacheSharedString)
            {
                if (!sharedStringsCache.TryGetValue(valueStr, out SharedStringCacheItem item))
                {
                    item = new SharedStringCacheItem { Position = sharedStringsCache.Count, Value = valueStr };
                    sharedStringsCache.Add(valueStr, item);
                }
                valueStr = item.Position.ToString();
                cellChildren[colIndex] = CreateHeaderCell(colIndex, rowIndex, valueStr, true);
                xlSharedStringsCount++;
            }
            else
            {
                cellChildren[colIndex] = CreateHeaderCell(colIndex, rowIndex, valueStr, false);
            }
        }
        headerRow.Append(cellChildren);
        return headerRow;
    }

    protected virtual Row CreateRowFromRecord(IDataRecord record, int rowIndex, bool preCacheSharedString = true)
    {
        Cell[] cellChildren = new Cell[WorksheetColumns.Count];
        var newRow = new Row { RowIndex = (uint)rowIndex + 1 };
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            string valueStr = GetValue(record[colIndex], columnInfo);

            if (preCacheSharedString && columnInfo.IsSharedString)
            {
                if (!sharedStringsCache.TryGetValue(valueStr, out SharedStringCacheItem item))
                {
                    item = new SharedStringCacheItem { Position = sharedStringsCache.Count, Value = valueStr };
                    sharedStringsCache.Add(valueStr, item);
                }
                valueStr = item.Position.ToString();
                cellChildren[colIndex] = CreateCell(colIndex, rowIndex, valueStr, columnInfo, true);
                xlSharedStringsCount++;
            }
            else
            {
                cellChildren[colIndex] = CreateCell(colIndex, rowIndex, valueStr, columnInfo, false);
            }
        }
        newRow.Append(cellChildren);
        __STATE = STATE_IMPORT;
        return newRow;
    }

    private SheetData CreateSheetData(SpreadsheetDocument document, Worksheet worksheet, DataTable dataTable)
    {
        var sheetData = new SheetData();

        int rowIndex = 0;
        int numOfRows = dataTable.Rows.Count;

        //TODO Switch to RowProxy ?
        //https://github.com/pre-alpha-final/openxml-memory-usage-hack/blob/master/docs/compare_code.png
        //https://github.com/dotnet/Open-XML-SDK/issues/807

        var headerRow = new Row { RowIndex = (uint)rowIndex + 1 };
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            Cell headerCell = CreateColumnHeader(colIndex, rowIndex, columnInfo);

            headerRow.AppendChild(headerCell);
        }
        sheetData.AppendChild(headerRow);
        _logger.LogTrace("{0} Columns in total", WorksheetColumns.Count);

        rowIndex = 1;
        var rowChildren = new List<Row>(numOfRows);
        var cellChildren = new Cell[WorksheetColumns.Count];
        foreach (DataRow dsrow in dataTable.Rows)
        {
            var newRow = new Row { RowIndex = (uint)rowIndex + 1 };
            for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
            {
                Cell dataCell = CreateCellFromDataType(colIndex, rowIndex, dsrow[colIndex]);
                cellChildren[colIndex] = dataCell;
            }
            newRow.Append(cellChildren);
            rowChildren.Add(newRow);
            rowIndex++;
        }
        sheetData.Append(rowChildren);
        _logger.LogTrace("{0} records with {1} columns has been added.", numOfRows, WorksheetColumns.Count);

        worksheet.AppendChild(sheetData);

        return sheetData;
    }

    protected SharedStringTablePart CreateSharedStringTablePart(SpreadsheetDocument document)
    {
        SharedStringTablePart sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(sharedStringPartRelationshipId);
        if (Context.CanUseRelativePaths)
        {
            sharedStringPartRelationshipId = document.UpdateWorkbookRelationshipsPath(sharedStringPart, ExcelConstants.SharedStringsRelationshipType);
        }
        return sharedStringPart;
    }

    protected virtual SharedStringTable CreateSharedStringTable(
        SpreadsheetDocument document, SharedStringTablePart sharedStringPart, Dictionary<string, SharedStringCacheItem> dict, int count)
    {
        var sharedStringTable = new SharedStringTable
        {
            UniqueCount = UInt32Value.FromUInt32((uint)dict.Count),
            Count = UInt32Value.FromUInt32((uint)count)
        };

        sharedStringPart.SharedStringTable = sharedStringTable;

        foreach (var item in sharedStringsCache.Values.OrderBy(i => i.Position))
        {
            var sharedStringItem = new SharedStringItem(
                new Text(item.Value));
            sharedStringTable.AppendChild(sharedStringItem);
        }

        return sharedStringTable;
    }


    protected virtual SharedStringTable CreateSharedStringTable(
        SpreadsheetDocument document, SharedStringTablePart sharedStringPart, DataTable dataTable)
    {
        SharedStringTable sharedStringTable = new SharedStringTable
        {
            UniqueCount = UInt32Value.FromUInt32(0U),
            Count = UInt32Value.FromUInt32(0U)
        };

        sharedStringPart.SharedStringTable = sharedStringTable;

        int count = 0;
        int uniqueCount = 0;

        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            if (!sharedStringsCache.ContainsKey(columnInfo.ColumnName))
            {
                sharedStringsCache.Add(columnInfo.ColumnName, 
                    new SharedStringCacheItem 
                    { 
                        Position = uniqueCount, 
                        Value = columnInfo.ColumnName 
                    });
                
                uniqueCount++;

                var sharedStringItem = new SharedStringItem(new Text(columnInfo.ColumnName));
                sharedStringTable.AppendChild(sharedStringItem);
            }
        }
        count += WorksheetColumns.Count;

        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            if (!columnInfo.IsSharedString)
                continue;

            foreach (DataRow dsrow in dataTable.Rows)
            {
                object val = dsrow[colIndex];

                if (val == DBNull.Value)
                    continue;

                string resultValue = GetValue(val, columnInfo);
                if (!sharedStringsCache.ContainsKey(resultValue))
                {
                    sharedStringsCache.Add(resultValue, new SharedStringCacheItem { Position = uniqueCount, Value = resultValue });
                    uniqueCount++;

                    var sharedStringItem = new SharedStringItem(new Text(resultValue));
                    sharedStringPart.SharedStringTable.AppendChild(sharedStringItem);
                }
            }

            count += dataTable.Rows.Count;
        }

        sharedStringTable.Count = (uint)count;
        sharedStringTable.UniqueCount = (uint)uniqueCount;
        sharedStringTable.Save();

        return sharedStringTable;
    }

    protected string GetSharedStringItem(string text)
    {
        return sharedStringsCache[text].Position.ToString();
    }

    protected Cell CreateStringCell(int columnIndex, int rowIndex, string cellValue, uint styleIndex = 0)
    {
        Cell cell = new Cell();
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "str"));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString()));
        cell.CellValue = new CellValue(cellValue);
        return cell;
    }

    protected Cell CreateValueCell(int columnIndex, int rowIndex, object cellValue, uint styleIndex = 0)
    {
        Cell cell = new Cell();
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "n"));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString()));
        cell.CellValue = new CellValue(cellValue.ToString());
        return cell;
    }

    protected Cell CreateDateCell(int columnIndex, int rowIndex, object cellValue, uint styleIndex = 0)
    {
        Cell cell = new Cell();
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "d"));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString()));
        cell.CellValue = new CellValue(cellValue.ToString());
        return cell;
    }

    protected Cell CreateInlineStringCell(int columnIndex, int rowIndex, string cellValue, uint styleIndex = 0)
    {
        Cell cell = new Cell();
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "inlineStr"));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString()));

        InlineString inlineString = new InlineString
        {
            Text = new Text
            {
                Text = cellValue
            }
        };

        cell.AppendChild(inlineString);
        return cell;
    }

    protected Cell CreateSharedStringCell(int columnIndex, int rowIndex, string cellValue, uint styleIndex, bool isValueSharedString = false)
    {
        Cell cell = new Cell();
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "t", string.Empty, "s"));
        cell.SetAttribute(new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString()));
        cell.CellValue = new CellValue(isValueSharedString ? cellValue : GetSharedStringItem(cellValue));
        return cell;
    }

    public static string GetCellReference(int columnIndex, int rowIndex)
    {
        return $"{WorksheetColumnInfo.GetColumnName(columnIndex)}{rowIndex + 1}";
    }

    protected Cell CreateColumnHeader(int columnIndex, int rowIndex, WorksheetColumnInfo columnInfo)
    {
        return CreateHeaderCell(columnIndex, rowIndex, columnInfo.Caption, false);
    }

    protected virtual Cell CreateHeaderCell(int columnIndex, int rowIndex, string value, bool isValueSharedString = false)
    {
        return CreateSharedStringCell(columnIndex, rowIndex, value, GetHeaderStyleIndex(), isValueSharedString);
    }

    private Cell CreateCell(int columnIndex, int rowIndex, string value, WorksheetColumnInfo columnInfo, bool isValueSharedString = false)
    {
        Cell dataCell = null;
        if (columnInfo.IsFloat)
        {
            if (columnInfo.IsSharedString)
            {
                dataCell = CreateSharedStringCell(columnIndex, rowIndex, value, GetDoubleStyleId(), isValueSharedString);
            }
            else
            {
                dataCell = CreateValueCell(columnIndex, rowIndex, value, GetDoubleStyleId());
            }
        }
        else if (columnInfo.IsDate)
        {

            if (Context.DateTimeAsString)
            {
                if (columnInfo.IsSharedString)
                {
                    dataCell = CreateSharedStringCell(columnIndex, rowIndex, value, GetDateStyleId(), isValueSharedString);
                }
                else if (columnInfo.IsInlineString)
                {
                    dataCell = CreateInlineStringCell(columnIndex, rowIndex, value, GetDateStyleId());
                }
                else
                {
                    dataCell = CreateStringCell(columnIndex, rowIndex, value, GetDateStyleId());
                }
            }
            else
            {
                dataCell = CreateDateCell(columnIndex, rowIndex, value, GetDateStyleId());
            }
        }
        else if (columnInfo.IsInteger)
        {
            if (columnInfo.IsSharedString)
            {
                dataCell = CreateSharedStringCell(columnIndex, rowIndex, value, GetIntegerStyleId(), isValueSharedString);
            }
            else
            {
                dataCell = CreateValueCell(columnIndex, rowIndex, value, GetIntegerStyleId());
            }
        }
        else
        {
            if (columnInfo.IsSharedString)
            {
                dataCell = CreateSharedStringCell(columnIndex, rowIndex, value, GetTextStyleId(), isValueSharedString);
            }
            else if (columnInfo.IsInlineString)
            {
                dataCell = CreateInlineStringCell(columnIndex, rowIndex, value, GetTextStyleId());
            }
            else
            {
                dataCell = CreateStringCell(columnIndex, rowIndex, value, GetTextStyleId());
            }
        }

        return dataCell;
    }

    protected Cell CreateCellFromDataType(int columnIndex, int rowIndex, object value)
    {
        var columnInfo = WorksheetColumns[columnIndex];
        string strValue = GetValue(value, columnInfo);
        return CreateCell(columnIndex, rowIndex, strValue, columnInfo);
    }

    //https://github.com/OfficeDev/Open-XML-SDK/issues/90
    protected string GetXMLWithDefaultNamespace(string outerXml, string defaultNamespace = ExcelConstants.DefaultSpreadsheetNamespace, string prefix = "x")
    {
        var xml = XDocument.Parse(outerXml);
        if (xml.Root != null)
        {
            RemoveNamespacePrefix(xml.Root, prefix);

            XNamespace xmlns = defaultNamespace;
            xml.Root.Name = xmlns + xml.Root.Name.LocalName;
        }

        if (Context.CanFixXmlDeclarations)
        {
            xml.Declaration = new XDeclaration("1.0", "UTF-8", "yes");
        }

        return xml.ToStringWithDeclaration().Replace(" xmlns=\"\"", string.Empty);
    }

    public static void RemoveNamespacePrefix(XElement element, string prefix)
    {
        if (element.Name.Namespace != null)
            element.Name = element.Name.LocalName;

        var attributes = element.Attributes().ToArray();
        element.RemoveAttributes();
        foreach (var attr in attributes)
        {
            var newAttr = attr;
            if (attr.Name.Namespace != null
                && attr.Name.Namespace.NamespaceName == XNamespace.Xmlns.NamespaceName
                && attr.Name.LocalName == prefix)
            {
                newAttr = new XAttribute("xmlns", attr.Value);
            }
            element.Add(newAttr);
        };

        foreach (var child in element.Descendants())
            RemoveNamespacePrefix(child, prefix);
    }

    protected ConnectionsPart CreateConnectionsPart(WorkbookPart workbookPart)
    {
        ConnectionsPart connectionsPart = workbookPart.AddNewPart<ConnectionsPart>();
        connectionsPart.Connections = new Connections();
        return connectionsPart;
    }

    protected Connection CreateConnection(
        WorksheetPart worksheetPart, WorkbookPart workbookPart, ConnectionsPart connectionsPart,
        string tableName, string connectionString, string sqlCommand)
    {
        Connection c = new Connection()
        {
            Id = 1,
            Name = tableName,
            //https://msdn.microsoft.com/en-us/library/office/ff839396.aspx
            Type = 1, //xlConnectionTypeOLEDB
            RefreshedVersion = 5,
            Background = true,
            SaveData = true,
            RefreshOnLoad = true,
            Credentials = CredentialsMethodValues.Integrated,
            SavePassword = false,
            DatabaseProperties = new DatabaseProperties()
            {
                Command = sqlCommand,
                //https://msdn.microsoft.com/en-us/library/office/ff197456.aspx
                CommandType = 2, //xlCmdSql
                Connection = connectionString
            },
            OnlyUseConnectionFile = false,
            KeepAlive = false
        };

        connectionsPart.Connections.Append(c);

        QueryTablePart qt = worksheetPart.AddNewPart<QueryTablePart>(tableName);
        qt.QueryTable = new QueryTable()
        {
            Name = tableName,
            ConnectionId = c.Id,
            AutoFormatId = 16,
            ApplyNumberFormats = true,
            ApplyBorderFormats = true,
            ApplyFontFormats = true,
            ApplyPatternFormats = true,
            ApplyAlignmentFormats = false,
            ApplyWidthHeightFormats = false,
            RefreshOnLoad = true
        };

        DefinedNames definedNames = new DefinedNames();
        DefinedName definedName = new DefinedName()
        {
            Name = tableName,
            Text = tableName + "!$A$1:$A$1",
            Description = tableName,
            Comment = tableName
        };

        definedNames.Append(definedName);
        workbookPart.Workbook.Append(definedNames);
        return c;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposed)
            return;

        if (disposing)
        {
            // Free any other managed objects here.
            Close();

            if (xlDocument != null)
            {
                xlDocument.Dispose();
                xlDocument = null;
            }
        }

        // Free any unmanaged objects here.
        //
        disposed = true;
    }

    ~ExcelExportAdapter()
    {
        Dispose(false);
    }

    protected uint GetDateStyleId()
    {
        return xlStylesPart?.DateStyleId ?? 0;
    }

    protected uint GetTextStyleId()
    {
        return xlStylesPart?.TextStyleId ?? 0;
    }

    protected uint GetIntegerStyleId()
    {
        return xlStylesPart?.IntegerStyleId ?? 0;
    }

    protected uint GetDoubleStyleId()
    {
        return xlStylesPart?.DoubleStyleId ?? 0;
    }

    protected uint GetHeaderStyleIndex()
    {
        return xlStylesPart?.HeaderStyleIndex ?? 0;
    }

}