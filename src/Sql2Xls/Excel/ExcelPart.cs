using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sql2Xls.Excel
{
    public abstract class ExcelPart
    {
        public SpreadsheetDocument Document { get; private set; }
        public string RelationshipId { get; protected set; }
        public ExcelExportContext Context { get; private set; }

        public ExcelPart(SpreadsheetDocument document, string relationshipId, ExcelExportContext context)
        {
            Document = document;
            RelationshipId = relationshipId;
            Context = context;
        }
    }
}
