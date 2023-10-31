using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Sql2Xls.Excel.Extensions;
using Sql2Xls.Excel.Parts;
using System.Data;

namespace Sql2Xls.Excel.Adapters;

public class ExcelExportSAXAdapter : ExcelExportAdapter
{
    private readonly ILogger<ExcelExportSAXAdapter> _logger;
    
    private readonly Cell _cellObject;
    private readonly InlineString _inlineStringObject;

    private OpenXmlWriter _xlWorksheetPartXmlWriter;

    public ExcelExportSAXAdapter(ILogger<ExcelExportSAXAdapter> logger) : base(logger)
    {
        _logger = logger;
        _cellObject = new Cell();
        _inlineStringObject = new InlineString();
    }

    public override SpreadsheetDocument Open()
    {
        xlDocument = SpreadsheetDocument.Create(Context.FileName, SpreadsheetDocumentType.Workbook);

        CreateExtendedFileProperties(xlDocument);
        CreateCoreFileProperties(xlDocument);

        xlWorkbookPart = CreateWorkbookPart(xlDocument);
        xlStylesPart = CreateWorkbookStylesPart(xlDocument, xlWorkbookPart);
        xlThemePart = CreateThemePart(xlDocument, xlWorkbookPart);
        xlSharedStringTablePart = CreateSharedStringTablePart(xlDocument);

        var sheetInfo = CreateSpreadSheetInfo();

        xlWorkbook = CreateWorkbook(xlWorkbookPart, sheetInfo);
        xlWorksheetPart = CreateWorksheetPart(xlWorkbookPart);

        _xlWorksheetPartXmlWriter = OpenXmlWriter.Create(xlWorksheetPart);
        CreateWorksheetPreSAX(xlDocument, xlWorksheetPart, _xlWorksheetPartXmlWriter);
        xlSheetData = new SheetData();
        _xlWorksheetPartXmlWriter.WriteStartElement(xlSheetData);

        __STATE = STATE_OPEN;

        return xlDocument;
    }

    protected override void AddHeaderRow(int rowIndex)
    {
        CreateHeaderRow(_xlWorksheetPartXmlWriter, rowIndex, true);
        __STATE = STATE_IMPORT;
    }

    private void CreateHeaderRow(OpenXmlWriter openXmlWriter, int rowIndex, bool preCacheSharedString = true)
    {
        openXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            string valueStr = columnInfo.Caption;

            if (preCacheSharedString)
            {
                if (!sharedStringsCache.TryGetValue(valueStr, out SharedStringCacheItem item))
                {
                    item = new SharedStringCacheItem { Position = sharedStringsCache.Count, Value = valueStr };
                    sharedStringsCache.Add(valueStr, item);
                }
                valueStr = item.Position.ToString();
                CreateColumnHeaderSAX(openXmlWriter, colIndex, rowIndex, valueStr, true);
                xlSharedStringsCount++;
            }
            else
            {
                CreateColumnHeaderSAX(openXmlWriter, colIndex, rowIndex, valueStr, false);
            }
        }
        openXmlWriter.WriteEndElement();
    }

    public override void AddDataRecord(IDataRecord dataRecord)
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

        CreateRowFromRecordSAX(_xlWorksheetPartXmlWriter, dataRecord, currentRow++, true);
    }

    private void CreateRowFromRecordSAX(OpenXmlWriter openXmlWriter, IDataRecord record, int rowIndex, bool preCacheSharedString = true)
    {
        openXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            string valueStr = GetValue(record[colIndex], columnInfo);

            if (preCacheSharedString)
            {
                if (!sharedStringsCache.TryGetValue(valueStr, out SharedStringCacheItem item))
                {
                    item = new SharedStringCacheItem { Position = sharedStringsCache.Count, Value = valueStr };
                    sharedStringsCache.Add(valueStr, item);
                }
                valueStr = item.Position.ToString();
                CreateCellSAX(openXmlWriter, colIndex, rowIndex, valueStr, WorksheetColumns[colIndex], true);
                xlSharedStringsCount++;
            }
            else
            {
                CreateCellSAX(openXmlWriter, colIndex, rowIndex, valueStr, WorksheetColumns[colIndex], false);
            }
        }
        openXmlWriter.WriteEndElement();

        __STATE = STATE_IMPORT;
    }

    public override void Close()
    {
        if (__STATE == STATE_OPEN || __STATE == STATE_IMPORT)
        {
            if (xlDocument != null)
            {
                if (_xlWorksheetPartXmlWriter != null)
                {
                    _xlWorksheetPartXmlWriter.WriteEndElement(); //SheedData
                }

                CreateSharedStringTable(xlDocument, xlSharedStringTablePart, sharedStringsCache, xlSharedStringsCount);

                if (_xlWorksheetPartXmlWriter != null)
                {
                    CreateWorksheetPostSAX(_xlWorksheetPartXmlWriter);
                    _xlWorksheetPartXmlWriter = null;
                }

                xlDocument.Dispose();
                xlDocument = null;
            }
            __STATE = STATE_CLOSED;
        }
    }

    protected override void CreateExtendedFileProperties(SpreadsheetDocument document)
    {
        if (!Context.CanCreateExtendedFileProperties)
            return;

        new ExcelExtendedFilePropertiesPart(document, extendedFilePropertiesPartRelationshipId, Context).CreateSAX();
    }

    protected override Workbook CreateWorkbook(WorkbookPart workbookPart, List<Sheet> sheetsInfo)
    {
        OpenXmlWriter openXmlWriter = OpenXmlWriter.Create(workbookPart);
        openXmlWriter.WriteStartDocument(true);
        var openXmlAttributes = new List<OpenXmlAttribute>();
        var namespaceDeclarations = new List<KeyValuePair<string, string>>
        {
            KeyValuePair.Create("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        };

        Workbook workbook = new Workbook();
        openXmlWriter.WriteStartElement(workbook, openXmlAttributes, namespaceDeclarations);

        openXmlWriter.WriteElement(new FileVersion
        {
            ApplicationName = StringValue.FromString("xl"),
            LastEdited = StringValue.FromString("6"),
            LowestEdited = StringValue.FromString("5"),
            BuildVersion = StringValue.FromString("14420")
            //CodeName = "{7A2D7E96-6E34-419A-AE5F-296B3A7E7977}" 
        });

        openXmlWriter.WriteElement(new WorkbookProperties
        {
            CodeName = "ThisWorkbook",
            DefaultThemeVersion = UInt32Value.FromUInt32(124226U)
        });

        openXmlWriter.WriteStartElement(new BookViews());
        openXmlWriter.WriteElement(new WorkbookView());
        openXmlWriter.WriteEndElement(); //BookViews

        openXmlWriter.WriteStartElement(new Sheets());

        foreach (var sheet in sheetsInfo)
        {
            openXmlWriter.WriteElement(sheet);
        }

        openXmlWriter.WriteEndElement(); //Sheets
        openXmlWriter.WriteEndElement(); //Workbook

        openXmlWriter.Close();

        return workbook;
    }

    private void CreateWorksheetPreSAX(SpreadsheetDocument document, WorksheetPart worksheetPart, OpenXmlWriter openXmlWriter)
    {        
        openXmlWriter.WriteStartDocument(true);

        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute("Ignorable", "mc", "x14ac xr xr2 xr3"),
            new OpenXmlAttribute("xr", "uid", ExcelConstants.SpreadsheetMlRev1, "{00000000-0001-0000-0000-000000000000}")
        };

        var namespaceDeclarations = new List<KeyValuePair<string, string>>
        {
            KeyValuePair.Create("r", ExcelConstants.RelationshipsNamespace),
            KeyValuePair.Create("mc", ExcelConstants.MarkupCompatibility),
            KeyValuePair.Create("x14ac", ExcelConstants.SpreadsheetMlAc),
            KeyValuePair.Create("xr", ExcelConstants.SpreadsheetMlRev1),
            KeyValuePair.Create("xr2", ExcelConstants.SpreadsheetMlRev2),
            KeyValuePair.Create("xr3", ExcelConstants.SpreadsheetMlRev3)
        };

        Worksheet worksheet = new Worksheet();
        openXmlWriter.WriteStartElement(worksheet, openXmlAttributes, namespaceDeclarations);

        if (Context.CanUseRelativePaths)
        {
            worksheetPartRelationshipId = document.UpdateWorkbookRelationshipsPath(worksheetPart, ExcelConstants.WorksheetRelationshipType);
        }

        /*
        openXmlWriter.WriteElement(
            new SheetDimension
            {
                Reference = StringValue.FromString(String.Format("A1:{0}", GetCellReference(dataTable.Columns.Count - 1, dataTable.Rows.Count + 1)))
            });
        */

        openXmlWriter.WriteStartElement(new SheetViews());
        openXmlWriter.WriteElement(
            new SheetView
            {
                TabSelected = BooleanValue.FromBoolean(true),
                WorkbookViewId = UInt32Value.FromUInt32(0U)
            });
        openXmlWriter.WriteEndElement();

        openXmlWriter.WriteElement(
            new SheetFormatProperties
            {
                DefaultRowHeight = DoubleValue.FromDouble(15D),
                DyDescent = DoubleValue.FromDouble(0.25D)
            });

        openXmlWriter.WriteStartElement(new Columns());
        openXmlWriter.WriteElement(
            new Column
            {
                Min = UInt32Value.FromUInt32(1U),
                Max = UInt32Value.FromUInt32((uint)WorksheetColumns.Count),
                Width = DoubleValue.FromDouble(20D),
                CustomWidth = BooleanValue.FromBoolean(true)
            });

        openXmlWriter.WriteEndElement();        
    }

    private void CreateWorksheetPostSAX(OpenXmlWriter openXmlWriter)
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

        openXmlWriter.WriteElement(pageMargins);

        HeaderFooter headerFooter = new HeaderFooter();
        openXmlWriter.WriteElement(headerFooter);
        openXmlWriter.WriteEndElement();        
    }

    protected override void CreateWorksheet(SpreadsheetDocument document, WorksheetPart worksheetPart, DataTable dataTable)
    {
        using var xmlWriter = OpenXmlWriter.Create(xlWorksheetPart);
        {
            CreateWorksheetPreSAX(document, worksheetPart, xmlWriter);
            CreateSheetDataSAX(xmlWriter, dataTable);
            CreateWorksheetPostSAX(xmlWriter);
            xmlWriter.Close();
        }        
    }

    private void CreateSheetDataSAX(OpenXmlWriter openXmlWriter, DataTable dataTable)
    {
        SheetData sheetData = new SheetData();
        openXmlWriter.WriteStartElement(sheetData);

        int rowIndex = 0;
        int numOfRows = dataTable.Rows.Count;

        openXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            CreateColumnHeaderSAX(openXmlWriter, colIndex, rowIndex, columnInfo.Caption, false);
        }
        openXmlWriter.WriteEndElement();

        rowIndex = 1;
        foreach (DataRow dsrow in dataTable.Rows)
        {
            openXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
            for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
            {
                CreateCellFromDataTypeSAX(openXmlWriter, colIndex, rowIndex, dsrow[colIndex]);
            }
            openXmlWriter.WriteEndElement();
            rowIndex++;
        }

        openXmlWriter.WriteEndElement();

        _logger.LogTrace("{0} records with {1} columns has been added.", numOfRows, WorksheetColumns.Count);
    }

    private void CreateColumnHeaderSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string caption, bool isValueSharedString = false)
    {
        CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, caption, GetHeaderStyleIndex(), isValueSharedString);
    }

    private void CreateCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, WorksheetColumnInfo columnInfo, bool isValueSharedString = false)
    {
        if (columnInfo.IsFloat)
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDoubleStyleId(), isValueSharedString);
            }
            else
            {
                CreateValueCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDoubleStyleId());
            }
        }
        else if (columnInfo.IsDate)
        {
            if (Context.DateTimeAsString)
            {
                if (columnInfo.IsSharedString)
                {
                    CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDateStyleId(), isValueSharedString);
                }
                else if (columnInfo.IsInlineString)
                {
                    CreateInlineStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDateStyleId());
                }
                else
                {
                    CreateStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDateStyleId());
                }
            }
            else
            {
                CreateDateCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetDateStyleId());
            }
        }
        else if (columnInfo.IsInteger)
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetIntegerStyleId(), isValueSharedString);
            }
            else
            {
                CreateValueCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetIntegerStyleId());
            }
        }
        else
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetTextStyleId(), isValueSharedString);
            }
            else if (columnInfo.IsInlineString)
            {
                CreateInlineStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetTextStyleId());
            }
            else
            {
                CreateStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, GetTextStyleId());
            }
        }
    }

    private void CreateCellFromDataTypeSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, object value)
    {
        var columnInfo = WorksheetColumns[columnIndex];
        string strValue = GetValue(value, columnInfo);
        CreateCellSAX(openXmlWriter, columnIndex, rowIndex, strValue, columnInfo, false);
    }

    private void CreateSharedStringCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0, bool isValueSharedString = false)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "s"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(isValueSharedString ? value : GetSharedStringItem(value)));
        openXmlWriter.WriteEndElement();
    }

    private void CreateInlineStringCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "inlineStr"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteStartElement(_inlineStringObject);
        openXmlWriter.WriteElement(new Text { Text = value });
        openXmlWriter.WriteEndElement();
        openXmlWriter.WriteEndElement();
    }

    private void CreateValueCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, object value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "n"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(new Cell(), openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(value.ToString()));
        openXmlWriter.WriteEndElement();
    }

    private void CreateDateCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "d"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(value));
        openXmlWriter.WriteEndElement();
    }

    private void CreateStringCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "str"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(value));
        openXmlWriter.WriteEndElement();
    }

    protected override SharedStringTable CreateSharedStringTable(SpreadsheetDocument document, SharedStringTablePart sharedStringPart, DataTable dataTable)
    {
        _logger.LogTrace("Creating Shared String Table");

        int count = 0;
        int uniqueCount = 0;

        for (int colIndex = 0; colIndex < WorksheetColumns.Count; colIndex++)
        {
            var columnInfo = WorksheetColumns[colIndex];
            if (!sharedStringsCache.ContainsKey(columnInfo.ColumnName))
            {
                sharedStringsCache.Add(columnInfo.ColumnName, new SharedStringCacheItem { Position = uniqueCount, Value = columnInfo.ColumnName });
                uniqueCount++;
            }
        }
        count += dataTable.Columns.Count;

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
                }
            }

            count += dataTable.Rows.Count;
        }

        return CreateSharedStringTable(document, sharedStringPart, sharedStringsCache, count);
    }

    protected override SharedStringTable CreateSharedStringTable(SpreadsheetDocument document, SharedStringTablePart sharedStringPart, Dictionary<string, SharedStringCacheItem> dict, int count)
    {
        SharedStringTable sharedStringTable = new SharedStringTable
        {
            UniqueCount = UInt32Value.FromUInt32((uint)count),
            Count = UInt32Value.FromUInt32((uint)dict.Count)
        };

        OpenXmlWriter openXmlWriter = OpenXmlWriter.Create(sharedStringPart);
        openXmlWriter.WriteStartDocument(true);
        openXmlWriter.WriteStartElement(sharedStringTable);
        foreach (var item in sharedStringsCache.Values.OrderBy(i => i.Position))
        {
            openXmlWriter.WriteStartElement(new SharedStringItem());
            openXmlWriter.WriteElement(new Text(item.Value));
            openXmlWriter.WriteEndElement();
        }
        openXmlWriter.WriteEndElement();
        openXmlWriter.Close();

        return sharedStringTable;
    }

}
