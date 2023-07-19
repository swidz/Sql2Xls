using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Sql2Xls.Excel.Extensions;

namespace Sql2Xls.Excel.Parts;

public class ExcelStylesPart : ExcelPart
{
    public UInt32Value IntegerStyleId { get; private set; }
    public UInt32Value DoubleStyleId { get; private set; }
    public UInt32Value DateStyleId { get; private set; }
    public UInt32Value TextStyleId { get; private set; }
    public UInt32Value HeaderStyleIndex { get; private set; }

    public const string stylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    public ExcelStylesPart(SpreadsheetDocument document, string relationshipId, ExcelExportContext context)
        : base(document, relationshipId, context)
    {
        IntegerStyleId = UInt32Value.FromUInt32(0U);
        DoubleStyleId = UInt32Value.FromUInt32(0U);
        DateStyleId = UInt32Value.FromUInt32(0U);
        TextStyleId = UInt32Value.FromUInt32(0U);
        HeaderStyleIndex = UInt32Value.FromUInt32(0U);
    }

    public WorkbookStylesPart CreateWorkbookStylesPart(WorkbookPart workbookPart)
    {
        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>(RelationshipId);
        stylesPart.Stylesheet = GenerateStylesheet();
        stylesPart.Stylesheet.Save();

        if (Context.CanUseRelativePaths)
        {
            RelationshipId = Document.UpdateWorkbookRelationshipsPath(stylesPart, stylesRelationshipType);
        }

        //http://www.lateral8.com/articles/openxml-format-excel-values.html
        IntegerStyleId = CreateCellFormat(stylesPart.Stylesheet, UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(1));
        DoubleStyleId = CreateCellFormat(stylesPart.Stylesheet, UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(4));
        DateStyleId = CreateCellFormat(stylesPart.Stylesheet, UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(14));
        TextStyleId = CreateCellFormat(stylesPart.Stylesheet, UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(49));

        UInt32Value headerFontIndex = CreateFont(stylesPart.Stylesheet, "Calibri", 11D, true, System.Drawing.Color.Black);
        //UInt32Value headerFillIndex = CreateFill(stylesPart.Stylesheet, System.Drawing.Color.Transparent);
        HeaderStyleIndex = CreateCellFormat(stylesPart.Stylesheet, headerFontIndex, UInt32Value.FromUInt32(0), UInt32Value.FromUInt32(0));

        return stylesPart;
    }

    protected UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value numberFormatId)
    {
        CellFormat cellFormat = new CellFormat();

        if (fontIndex != null)
            cellFormat.FontId = fontIndex;

        if (fillIndex != null)
            cellFormat.FillId = fillIndex;

        if (numberFormatId != null)
        {
            cellFormat.NumberFormatId = numberFormatId;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        }

        styleSheet.CellFormats.Append(cellFormat);

        if (styleSheet.CellFormats.Count == null)
            styleSheet.CellFormats.Count = UInt32Value.FromUInt32(1U);
        else
            styleSheet.CellFormats.Count++;

        return styleSheet.CellFormats.Count - 1;
    }


    protected Stylesheet GenerateStylesheet()
    {
        NumberingFormats numFmts = new NumberingFormats { Count = UInt32Value.FromUInt32(0U) };

        FontName fontName0 = new FontName { Val = "Calibri" };
        FontSize sz0 = new FontSize
        {
            Val = new DoubleValue { Value = 11D }
        };
        Font font0 = new Font(sz0, fontName0);

        Fonts fonts = new Fonts();
        fonts.Append(font0);
        fonts.Count = UInt32Value.FromUInt32(1U);

        Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }) // Index 1 - default
            )
        {
            Count = UInt32Value.FromUInt32(2U)
        };

        var left0 = new LeftBorder();
        var right0 = new RightBorder();
        var top0 = new TopBorder();
        var bottom0 = new BottomBorder();
        var diagonal0 = new DiagonalBorder();

        var border0 = new Border(left0, right0, top0, bottom0, diagonal0);
        Borders borders = new Borders(border0)
        {
            Count = UInt32Value.FromUInt32(1U)
        };

        CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)1U };
        CellFormat cellFormat = new CellFormat()
        {
            NumberFormatId = UInt32Value.FromUInt32(0),
            FontId = UInt32Value.FromUInt32(0),
            BorderId = UInt32Value.FromUInt32(0),
            FillId = UInt32Value.FromUInt32(0)
        };
        cellStyleFormats.Append(cellFormat);

        CellFormats cellFormats = new CellFormats(
                new CellFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(0),
                    FontId = UInt32Value.FromUInt32(0),
                    FormatId = UInt32Value.FromUInt32(0),
                    ApplyNumberFormat = BooleanValue.FromBoolean(true),
                    ApplyFont = BooleanValue.FromBoolean(true),
                    ApplyProtection = BooleanValue.FromBoolean(true)
                }
            )
        {
            Count = UInt32Value.FromUInt32(1U)
        };

        CellStyles cellStyles = new CellStyles { Count = UInt32Value.FromUInt32(1U) };
        CellStyle cellStyle = new CellStyle { Name = StringValue.FromString("Normal"), FormatId = UInt32Value.FromUInt32(0U), BuiltinId = UInt32Value.FromUInt32(0U) };
        cellStyles.Append(cellStyle);

        DifferentialFormats dxfs = new DifferentialFormats { Count = UInt32Value.FromUInt32(0U) };

        Stylesheet styleSheet = new Stylesheet(numFmts, fonts, fills, borders, cellStyleFormats, cellFormats, cellStyles, dxfs);
        styleSheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        styleSheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

        return styleSheet;
    }

    protected UInt32Value CreateFont(Stylesheet styleSheet, string fontName, double? fontSize,
        bool isBold, System.Drawing.Color foreColor)
    {
        Font font = new Font();

        if (isBold == true)
        {
            Bold bold = new Bold();
            font.Append(bold);
        }

        if (fontSize.HasValue)
        {
            FontSize size = new FontSize()
            {
                Val = new DoubleValue() { Value = fontSize.Value }
            };

            font.Append(size);
        }


        Color color = new Color()
        {
            Rgb = new HexBinaryValue()
            {
                Value = HexBinaryValue.FromString(ColorHexConverter(foreColor).Replace("#", string.Empty))
            }
        };
        font.Append(color);

        if (!string.IsNullOrEmpty(fontName))
        {
            FontName name = new FontName()
            {
                Val = fontName
            };
            font.Append(name);
        }

        styleSheet.Fonts.Append(font);

        if (styleSheet.Fonts.Count == null)
        {
            styleSheet.Fonts.Count = UInt32Value.FromUInt32(1U);
        }
        else
        {
            styleSheet.Fonts.Count++;
        }

        return styleSheet.Fonts.Count - 1;
    }

    protected UInt32Value CreateFill(Stylesheet styleSheet, System.Drawing.Color fillColor)
    {
        Fill fill = new Fill(
            new PatternFill(
                new ForegroundColor()
                {
                    Rgb = new HexBinaryValue()
                    {
                        Value = HexBinaryValue.FromString(ColorHexConverter(fillColor).Replace("#", string.Empty))
                    }
                })
            {
                PatternType = PatternValues.Solid
            });

        styleSheet.Fills.Append(fill);
        if (styleSheet.Fills.Count == null)
        {
            styleSheet.Fills.Count = (UInt32Value)1U;
        }
        else
        {
            styleSheet.Fills.Count++;
        }

        return styleSheet.Fills.Count - 1;
    }

    protected string ColorHexConverter(System.Drawing.Color c)
    {
        //https://code-examples.net/en/q/248d2e
        return string.Format("#{0}{1}{2}{3}"
        , c.A.ToString("X").Length == 1 ? string.Format("0{0}", c.A.ToString("X")) : c.A.ToString("X")
        , c.R.ToString("X").Length == 1 ? string.Format("0{0}", c.R.ToString("X")) : c.R.ToString("X")
        , c.G.ToString("X").Length == 1 ? string.Format("0{0}", c.G.ToString("X")) : c.G.ToString("X")
        , c.B.ToString("X").Length == 1 ? string.Format("0{0}", c.B.ToString("X")) : c.B.ToString("X"));
    }

}
