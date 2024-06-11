using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Sql2Xls.Excel.Extensions;
using Sql2Xls.Excel.Parts;
using Sql2Xls.Extensions;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Sql2Xls.Excel.Adapters;

public class ExcelExportSAXAdapterV2 : IDisposable
{
    private readonly ILogger<ExcelExportSAXAdapterV2> _logger;
    private WorksheetColumnCollection _worksheetColumns;

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

    private readonly Dictionary<string, SharedStringCacheItem> _sharedStringsCache;

    private string _workbookPartRelationshipId = "rId1";
    private string _coreFilePropertiesPartRelationshipId = "rId2";
    private string _extendedFilePropertiesPartRelationshipId = "rId3";
    private string _worksheetPartRelationshipId = "rId1";
    private string _themePartRelationshipId = "rId2";
    private string _workbookStylesPartRelationshipId = "rId3";
    private string _sharedStringPartRelationshipId = "rId4";
    
    private uint _integerStyleId;
    private uint _doubleStyleId ;
    private uint _dateStyleId;
    private uint _textStyleId;
    private uint _headerStyleIndex;
    private bool _disposedValue;

    private readonly Cell _cellObject;
    private readonly InlineString _inlineStringObject;
    private readonly SharedStringItem _sharedStringItem;


    public ExcelExportSAXAdapterV2(ILogger<ExcelExportSAXAdapterV2> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _sharedStringsCache = new(10000);

        _cellObject = new Cell();
        _inlineStringObject = new InlineString();
        _sharedStringItem = new SharedStringItem();
    }

    public void LoadFromDataTable(DataTable dt)
    {        
        _worksheetColumns = WorksheetColumnCollection.Create(dt, Context);

        using var document = SpreadsheetDocument.Create(Context.FileName, SpreadsheetDocumentType.Workbook);        
        
        CreateExtendedProperties(document);
        CreateCoreFileProperties(document);

        //var workbookPart = CreateWorkbookPart(document);
        var workbookPart = document.AddWorkbookPart();
        document.ChangeIdOfPart(workbookPart, _workbookPartRelationshipId);

        if (Context.CanUseRelativePaths)
        {
            _workbookPartRelationshipId = document.UpdateDocumentRelationshipsPath(workbookPart, ExcelConstants.OfficeDocumentRelationshipType);
        }

        //var stylesPart = CreateWorkbookStylesPart(document, workbookPart);
        var stylesPart = new ExcelStylesPart(document, _workbookStylesPartRelationshipId, Context);
        stylesPart.CreateWorkbookStylesPart(workbookPart);
        _integerStyleId = stylesPart.IntegerStyleId;
        _doubleStyleId = stylesPart.DoubleStyleId;
        _dateStyleId = stylesPart.DateStyleId;
        _textStyleId = stylesPart.TextStyleId;
        _headerStyleIndex = stylesPart.HeaderStyleIndex;        

        //var themePart = CreateThemePart(document, workbookPart);
        if (Context.CanCreateThemePart)
        {
            ExcelThemePart themePart = new ExcelThemePart(document, _themePartRelationshipId, Context);
            themePart.CreateThemePart(workbookPart);
        }
        
        //var sharedStringTablePart = CreateSharedStringTablePart(document);
        SharedStringTablePart sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>(_sharedStringPartRelationshipId);
        if (Context.CanUseRelativePaths)
        {
            _sharedStringPartRelationshipId = document.UpdateWorkbookRelationshipsPath(sharedStringPart, ExcelConstants.SharedStringsRelationshipType);
        }

        //CreateSharedStringTable(document, sharedStringTablePart, dataTable);                            
        int count = 0;
        for (int colIndex = 0; colIndex < _worksheetColumns.Count; colIndex++)
        {
            var columnInfo = _worksheetColumns[colIndex];
            if (!_sharedStringsCache.ContainsKey(columnInfo.ColumnName))
            {
                _sharedStringsCache.Add(columnInfo.ColumnName, new SharedStringCacheItem
                {
                    Position = _sharedStringsCache.Count,
                    Value = columnInfo.ColumnName
                });
            }
        }
        count += _worksheetColumns.Count;

        for (int colIndex = 0; colIndex < _worksheetColumns.Count; colIndex++)
        {
            var columnInfo = _worksheetColumns[colIndex];
            if (!columnInfo.IsSharedString)
                continue;            

            foreach (DataRow dsrow in dt.Rows)
            {
                object val = dsrow[colIndex];

                if (val is null || val == DBNull.Value)
                    continue;

                var resultValue = columnInfo.GetStringValue(val);
                if (!_sharedStringsCache.ContainsKey(resultValue))
                {
                    _sharedStringsCache.Add(resultValue, new SharedStringCacheItem 
                    { 
                        Position = _sharedStringsCache.Count, 
                        Value = resultValue 
                    });
                }
            }
            count += dt.Rows.Count;
        }

        var sharedStringTable = new SharedStringTable
        {
            UniqueCount = UInt32Value.FromUInt32((uint)_sharedStringsCache.Count),
            Count = UInt32Value.FromUInt32((uint)count)
        };

        using var sharedStringXmlWriter = OpenXmlWriter.Create(sharedStringPart);
        {
            sharedStringXmlWriter.WriteStartDocument(true);
            sharedStringXmlWriter.WriteStartElement(sharedStringTable);
            foreach (var item in _sharedStringsCache.Values.OrderBy(i => i.Position))
            {
                sharedStringXmlWriter.WriteStartElement(_sharedStringItem);

                //System.IO.IOException: 'Stream was too long.'
                sharedStringXmlWriter.WriteElement(new Text(item.Value));
                sharedStringXmlWriter.WriteEndElement();                
            }
            sharedStringXmlWriter.WriteEndElement();            
        }

        //var sheetInfo = CreateSpreadSheetInfo();
        var sheetsInfo = new List<Sheet>(1)
        {
            new()
            {
                Name = StringValue.FromString(Context.SheetName),
                SheetId = UInt32Value.FromUInt32(1U),
                Id = _workbookPartRelationshipId
            }
        };


        //var workBook = CreateWorkbook(workbookPart, sheetInfo);
        using var openXmlWriter = OpenXmlWriter.Create(workbookPart);
        {
            openXmlWriter.WriteStartDocument(true);

            var openXmlAttributes = new List<OpenXmlAttribute>(0);
            var namespaceDeclarations = new List<KeyValuePair<string, string>>(1)
            {
                KeyValuePair.Create("r", ExcelConstants.RelationshipsNamespace)
            };

            var workbook = new Workbook();


            /*
            if (!String.IsNullOrEmpty(Context.Password))
            {
                const int spinCount = 100000;
                var saltValue = Guid.NewGuid().ToByteArray();
                var hash = HashHelper.ComputePasswordHash(Context.Password, saltValue, spinCount);

                workbook.WorkbookProtection = new WorkbookProtection()
                {
                    WorkbookPassword = new HexBinaryValue(HashHelper.HexPasswordConversion(Context.Password)),
                    WorkbookSaltValue = new Base64BinaryValue(Convert.ToBase64String(saltValue)),
                    WorkbookSpinCount = new UInt32Value((uint)spinCount),
                    WorkbookHashValue = new Base64BinaryValue(Convert.ToBase64String(hash)),
                    WorkbookAlgorithmName = new StringValue("SHA-512"),                                       
                };                
            }
            */
                       
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
        }

        //var worksheetPart = CreateWorksheetPart(workbookPart);
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(_worksheetPartRelationshipId);


        //CreateWorksheet(document, worksheetPart, dataTable);

        /*
        using var xmlWriter = OpenXmlWriter.Create(xlWorksheetPart);
        {
            CreateWorksheetPreSAX(document, worksheetPart, xmlWriter);
            CreateSheetDataSAX(xmlWriter, dataTable);
            CreateWorksheetPostSAX(xmlWriter);
            xmlWriter.Close();
        }  
        */


        using var worksheetXmlWriter = OpenXmlWriter.Create(worksheetPart);
        {
            //Pre
            worksheetXmlWriter.WriteStartDocument(true); //XML declaration

            
            var openXmlAttributes = new List<OpenXmlAttribute>(2)
            {
                new ("Ignorable", "mc", "x14ac xr xr2 xr3"),
                new ("xr", "uid", ExcelConstants.SpreadsheetMlRev1, "{00000000-0001-0000-0000-000000000000}")
            };

            var namespaceDeclarations = new List<KeyValuePair<string, string>>(7)
            {
                KeyValuePair.Create("", ExcelConstants.DefaultSpreadsheetNamespace),
                KeyValuePair.Create("r", ExcelConstants.RelationshipsNamespace),
                KeyValuePair.Create("mc", ExcelConstants.MarkupCompatibility),
                KeyValuePair.Create("x14ac", ExcelConstants.SpreadsheetMlAc),
                KeyValuePair.Create("xr", ExcelConstants.SpreadsheetMlRev1),
                KeyValuePair.Create("xr2", ExcelConstants.SpreadsheetMlRev2),
                KeyValuePair.Create("xr3", ExcelConstants.SpreadsheetMlRev3)
            };
            
            worksheetXmlWriter.WriteStartElement(new Worksheet(), openXmlAttributes, namespaceDeclarations);                       

            if (Context.CanUseRelativePaths)
            {
                _worksheetPartRelationshipId = document.UpdateWorkbookRelationshipsPath(worksheetPart, ExcelConstants.WorksheetRelationshipType);
            }            

            worksheetXmlWriter.WriteStartElement(new SheetViews());
            worksheetXmlWriter.WriteElement(
                new SheetView
                {
                    TabSelected = BooleanValue.FromBoolean(true),
                    WorkbookViewId = UInt32Value.FromUInt32(0U)
                });
            worksheetXmlWriter.WriteEndElement(); //SheetViews

            worksheetXmlWriter.WriteElement(
                new SheetFormatProperties
                {
                    DefaultRowHeight = DoubleValue.FromDouble(15D),
                    DyDescent = DoubleValue.FromDouble(0.25D)
                });

            worksheetXmlWriter.WriteStartElement(new Columns());
            worksheetXmlWriter.WriteElement(
                new Column
                {
                    Min = UInt32Value.FromUInt32(1U),
                    Max = UInt32Value.FromUInt32((uint)_worksheetColumns.Count),
                    Width = DoubleValue.FromDouble(20D),
                    CustomWidth = BooleanValue.FromBoolean(true)
                });
            worksheetXmlWriter.WriteEndElement(); //Columns


            //Main            
            worksheetXmlWriter.WriteStartElement(new SheetData());

            int rowIndex = 0;           
            worksheetXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
            for (int colIndex = 0; colIndex < _worksheetColumns.Count; colIndex++)
            {                
                CreateSharedStringCellSAX(worksheetXmlWriter, colIndex, rowIndex, _worksheetColumns[colIndex].Caption, _headerStyleIndex);
            }
            worksheetXmlWriter.WriteEndElement(); //Row

            rowIndex = 1;
            foreach (DataRow dsrow in dt.Rows)
            {
                worksheetXmlWriter.WriteStartElement(new Row { RowIndex = (uint)rowIndex + 1 });
                for (int colIndex = 0; colIndex < _worksheetColumns.Count; colIndex++)
                {
                    CreateCellFromDataTypeSAX(worksheetXmlWriter, colIndex, rowIndex, dsrow[colIndex]);
                }
                worksheetXmlWriter.WriteEndElement(); //Row
                rowIndex++;
            }

            worksheetXmlWriter.WriteEndElement(); //SheetData

            /*
            if(!String.IsNullOrEmpty(Context.Password))
            {
                const int spinCount = 100000;
                var saltValue = Guid.NewGuid().ToByteArray();
                var hash = HashHelper.ComputePasswordHash(Context.Password, saltValue, spinCount);

                worksheetXmlWriter.WriteElement(new SheetProtection()
                {
                    AlgorithmName = "SHA-512",
                    HashValue = Convert.ToBase64String(hash),
                    SaltValue = Convert.ToBase64String(saltValue),
                    SpinCount = UInt32Value.FromUInt32(spinCount),
                    Sheet = BooleanValue.FromBoolean(true),
                    Objects = BooleanValue.FromBoolean(true),
                    AutoFilter = BooleanValue.FromBoolean(true),
                    DeleteColumns = BooleanValue.FromBoolean(true),
                    DeleteRows = BooleanValue.FromBoolean(true),
                    FormatCells = BooleanValue.FromBoolean(true),
                    FormatColumns = BooleanValue.FromBoolean(true),
                    FormatRows = BooleanValue.FromBoolean(true),
                    InsertColumns = BooleanValue.FromBoolean(true),
                    InsertRows = BooleanValue.FromBoolean(true),
                    InsertHyperlinks = BooleanValue.FromBoolean(true),                    
                    Password = HashHelper.HexPasswordConversion(Context.Password),                    
                    PivotTables = BooleanValue.FromBoolean(true),
                    Scenarios = BooleanValue.FromBoolean(true),
                    SelectLockedCells = BooleanValue.FromBoolean(false),
                    SelectUnlockedCells = BooleanValue.FromBoolean(false),
                    Sort = BooleanValue.FromBoolean(true)                                        
                });                
            }
            */

            //Post           
            worksheetXmlWriter.WriteElement(new PageMargins()
            {
                Left = DoubleValue.FromDouble(0.7D),
                Right = DoubleValue.FromDouble(0.7D),
                Top = DoubleValue.FromDouble(0.75D),
                Bottom = DoubleValue.FromDouble(0.75D),
                Header = DoubleValue.FromDouble(0.3D),
                Footer = DoubleValue.FromDouble(0.3D)
            });

            worksheetXmlWriter.WriteElement(new HeaderFooter());




            worksheetXmlWriter.WriteEndElement(); //Worksheet         
        }

        if (Context.CanFixContentTypes || Context.CanFixXmlDeclarations || Context.CanRemoveAliasFromDefaultNamespace)
        {
            using (var archive = ZipFile.Open(Context.FileName, ZipArchiveMode.Update))
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

    private void CreateExtendedProperties(SpreadsheetDocument document)
    {
        if (!Context.CanCreateExtendedFileProperties)
            return;

        var extendedFilePropertiesPart = document.AddNewPart<ExtendedFilePropertiesPart>(_extendedFilePropertiesPartRelationshipId);

        OpenXmlWriter openXmlWriter = OpenXmlWriter.Create(extendedFilePropertiesPart);
        openXmlWriter.WriteStartDocument(true);
        openXmlWriter.WriteStartElement(new DocumentFormat.OpenXml.ExtendedProperties.Properties());
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.Application("Microsoft Excel"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.DocumentSecurity("0"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion("14.0000"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.SharedDocument("false"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.ScaleCrop("false"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.HyperlinksChanged("false"));
        openXmlWriter.WriteElement(new DocumentFormat.OpenXml.ExtendedProperties.LinksUpToDate("false"));
        openXmlWriter.WriteEndElement();
        openXmlWriter.Close();

        if (Context.CanUseRelativePaths)
        {
            document.UpdateDocumentRelationshipsPath(extendedFilePropertiesPart, ExcelConstants.ExtendedPropertiesRelationshipType);
        }
    }

    private void CreateCoreFileProperties(SpreadsheetDocument document)
    {
        if (!Context.CanCreateCoreFileProperties)
            return;

        CoreFilePropertiesPart part = document.AddNewPart<CoreFilePropertiesPart>(_coreFilePropertiesPartRelationshipId);
        using (var writer = new XmlTextWriter(part.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
        {
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.WriteRaw(Environment.NewLine);
            writer.WriteRaw("<cp:coreProperties ");
            writer.WriteRaw("xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
            writer.WriteRaw("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" ");
            writer.WriteRaw("xmlns:dcterms=\"http://purl.org/dc/terms/\" ");
            writer.WriteRaw("xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ");
            writer.WriteRaw("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            writer.WriteRaw("<dc:creator>" + Environment.UserName + "</dc:creator>");
            writer.WriteRaw("<cp:lastModifiedBy>" + Environment.UserName + "</cp:lastModifiedBy>");
            string createdDateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
            writer.WriteRaw("<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + createdDateTime + "</dcterms:created>");
            writer.WriteRaw("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + createdDateTime + "</dcterms:modified>");
            writer.WriteRaw("</cp:coreProperties>");
            writer.Flush();
        }

        if (Context.CanUseRelativePaths)
        {
            document.UpdateDocumentRelationshipsPath(part, ExcelConstants.CoreFilePropertiesRelationshipType);
        }
    }        

    private void CreateSharedStringCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>(3)
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "s"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };        

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(_sharedStringsCache[value].Position));
        openXmlWriter.WriteEndElement();
    }

    public static string GetCellReference(int columnIndex, int rowIndex)
    {
        return $"{WorksheetColumnInfo.GetColumnName(columnIndex)}{rowIndex + 1}";
    }        

    private void CreateCellFromDataTypeSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, object value)
    {
        var columnInfo = _worksheetColumns[columnIndex];
        string strValue = columnInfo.GetStringValue(value);
        CreateCellSAX(openXmlWriter, columnIndex, rowIndex, strValue, columnInfo);
    }

    private void CreateCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, WorksheetColumnInfo columnInfo)
    {                               
        if (columnInfo.IsFloat)
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _doubleStyleId);
            }
            else
            {
                CreateValueCellSAX(openXmlWriter, columnIndex, rowIndex, value, _doubleStyleId);
            }
        }
        else if (columnInfo.IsDateTime)
        {
            if (Context.DateTimeAsString)
            {
                if (columnInfo.IsSharedString)
                {
                    CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _dateStyleId);
                }
                else if (columnInfo.IsInlineString)
                {
                    CreateInlineStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _dateStyleId);
                }
                else
                {
                    CreateStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _dateStyleId);
                }
            }
            else
            {
                CreateDateCellSAX(openXmlWriter, columnIndex, rowIndex, value, _dateStyleId);
            }
        }
        else if (columnInfo.IsInteger)
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _integerStyleId);
            }
            else
            {
                CreateValueCellSAX(openXmlWriter, columnIndex, rowIndex, value, _integerStyleId);
            }
        }
        else
        {
            if (columnInfo.IsSharedString)
            {
                CreateSharedStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _textStyleId);
            }
            else if (columnInfo.IsInlineString)
            {
                CreateInlineStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _textStyleId);
            }
            else
            {
                CreateStringCellSAX(openXmlWriter, columnIndex, rowIndex, value, _textStyleId);
            }
        }
    }

    private void CreateInlineStringCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>(3)
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
        var openXmlAttributes = new List<OpenXmlAttribute>(3)
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "n"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(value.ToString()));
        openXmlWriter.WriteEndElement();
    }

    private void CreateDateCellSAX(OpenXmlWriter openXmlWriter, int columnIndex, int rowIndex, string value, uint styleIndex = 0)
    {
        var openXmlAttributes = new List<OpenXmlAttribute>(3)
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
        var openXmlAttributes = new List<OpenXmlAttribute>(3)
        {
            new OpenXmlAttribute(string.Empty, "r", string.Empty, GetCellReference(columnIndex, rowIndex)),
            new OpenXmlAttribute(string.Empty, "t", string.Empty, "str"),
            new OpenXmlAttribute(string.Empty, "s", string.Empty, styleIndex.ToString())
        };

        openXmlWriter.WriteStartElement(_cellObject, openXmlAttributes);
        openXmlWriter.WriteElement(new CellValue(value));        
        openXmlWriter.WriteEndElement();
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

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            _disposedValue = true;
        }
    }

    public void Dispose()
    {        
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
