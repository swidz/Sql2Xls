using DocumentFormat.OpenXml.Packaging;
using System.Xml;

namespace Sql2Xls.Excel;

public class ExcelCoreFilePropertiesPart : ExcelPart
{
    protected const string CoreFilePropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/core-properties";

    public ExcelCoreFilePropertiesPart(SpreadsheetDocument document, string relationshipId, ExcelExportContext context)
        : base(document, relationshipId, context)
    {
    }

    public void CreateDOM()
    {
        CoreFilePropertiesPart part = Document.AddNewPart<CoreFilePropertiesPart>(RelationshipId);
        using (XmlTextWriter writer = new XmlTextWriter(
                part.GetStream(FileMode.Create),
                System.Text.Encoding.UTF8))
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
            ExcelHelper.UpdateDocumentRelationshipsPath(Document, part, CoreFilePropertiesRelationshipType);
        }
    }
}
