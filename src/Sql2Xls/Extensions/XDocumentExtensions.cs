using System;
using System.IO;
using System.Text;
using System.Xml.Linq;

namespace Sql2Xls.Extensions;

public static class XDocumentExtensions
{
    public static string ToStringWithDeclaration(this XDocument doc)
    {
        if (doc == null)
        {
            throw new ArgumentNullException("doc");
        }
        var sb = new StringBuilder();
        using (TextWriter writer = new Utf8StringWriter(sb))
        {
            doc.Save(writer, SaveOptions.DisableFormatting);
        }
        return sb.ToString();
    }
}

//https://stackoverflow.com/questions/955611/xmlwriter-to-write-to-a-string-instead-of-to-a-file/955698#955698
public class Utf8StringWriter : StringWriter
{
    public Utf8StringWriter(StringBuilder sb)
        : base(sb)
    {
    }

    public override Encoding Encoding
    {
        get { return Encoding.UTF8; }
    }
}
