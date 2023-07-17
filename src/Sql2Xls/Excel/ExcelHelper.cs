using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;

namespace Sql2Xls.Excel;

public static class ExcelHelper
{
    //https://www.developerfusion.com/article/84390/programming-office-documents-with-open-xml/
    //https://stackoverflow.com/questions/10929054/openxml-spreadsheet-created-in-net-wont-open-in-ipad/15524301

    public static string UpdateWorkbookRelationshipsPath(SpreadsheetDocument document, OpenXmlPart part, string relationshipType)
    {
        string origPartRelationshipId = document.WorkbookPart.GetIdOfPart(part);
        string newPartRelationshipId = origPartRelationshipId;

        if (part.Uri.OriginalString.StartsWith("/xl/"))
        {
            var workbookpart = document.GetPackage().GetPart(document.WorkbookPart.Uri) as PackagePart;
            workbookpart?.DeleteRelationship(origPartRelationshipId);

            workbookpart = document.GetPackage().GetPart(document.WorkbookPart.Uri) as PackagePart;
            workbookpart?.CreateRelationship(new Uri(part.Uri.OriginalString.Replace("/xl/", string.Empty).Trim(), UriKind.Relative), TargetMode.Internal, relationshipType, origPartRelationshipId);

            var styleRelationships = document.GetPackage()
                .GetPart(document.WorkbookPart.Uri)
                .Relationships
                .Where(x => x.RelationshipType == relationshipType);
                
            //.GetRelationshipsByType(relationshipType);
            
            newPartRelationshipId = styleRelationships
                .Where(f => f.TargetUri.OriginalString == part.Uri.OriginalString.Replace("/xl/", string.Empty).Trim())
                .Single().Id;
        }

        return newPartRelationshipId;
    }

    
    public static string UpdateDocumentRelationshipsPath(SpreadsheetDocument document, OpenXmlPart part, string relationshipType)
    {
        string origPartRelationshipId = document.GetIdOfPart(part);
        string newPartRelationshipId = origPartRelationshipId;

        if (part.Uri.OriginalString.StartsWith("/"))
        {
            var package = document.GetPackage() as Package;
            package?.DeleteRelationship(origPartRelationshipId);

            package = document.GetPackage() as Package;
            package?.CreateRelationship(new Uri(part.Uri.OriginalString[1..].Trim(), UriKind.Relative), TargetMode.Internal, relationshipType, origPartRelationshipId);

            var styleRelationships = document.GetPackage()
                .Relationships
                .Where(x => x.RelationshipType == relationshipType);

            //.GetRelationshipsByType(relationshipType);

            newPartRelationshipId = styleRelationships
                .Where(f => f.TargetUri.OriginalString == part.Uri.OriginalString[1..].Trim())
                .Single().Id;
        }

        return newPartRelationshipId;
    }

}
