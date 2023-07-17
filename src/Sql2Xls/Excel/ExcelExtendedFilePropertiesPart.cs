﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sql2Xls.Excel
{
    public class ExcelExtendedFilePropertiesPart : ExcelPart
    {
        public const string ExtendedPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

        public ExcelExtendedFilePropertiesPart(SpreadsheetDocument document, string relationshipId, ExcelExportContext context)
            : base(document, relationshipId, context)
        {
        }

        public void CreateSAX()
        {
            var extendedFilePropertiesPart = Document.AddNewPart<ExtendedFilePropertiesPart>(RelationshipId);

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
                ExcelHelper.UpdateDocumentRelationshipsPath(Document, extendedFilePropertiesPart, ExtendedPropertiesRelationshipType);
            }
        }

        public void CreateDOM()
        {
            Document.AddNewPart<ExtendedFilePropertiesPart>(RelationshipId);
            Document.ExtendedFilePropertiesPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties
            {
                Application = new DocumentFormat.OpenXml.ExtendedProperties.Application("Microsoft Excel"),
                ApplicationVersion = new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion("14.0000"),
                DocumentSecurity = new DocumentFormat.OpenXml.ExtendedProperties.DocumentSecurity("0"),
                SharedDocument = new DocumentFormat.OpenXml.ExtendedProperties.SharedDocument("false"),
                ScaleCrop = new DocumentFormat.OpenXml.ExtendedProperties.ScaleCrop("false"),
                HyperlinksChanged = new DocumentFormat.OpenXml.ExtendedProperties.HyperlinksChanged("false"),
                LinksUpToDate = new DocumentFormat.OpenXml.ExtendedProperties.LinksUpToDate("false")
            };
            Document.ExtendedFilePropertiesPart.Properties.Save();

            if (Context.CanUseRelativePaths)
            {
                ExcelHelper.UpdateDocumentRelationshipsPath(Document, Document.ExtendedFilePropertiesPart, ExtendedPropertiesRelationshipType);
            }
        }
    }
}
