using Microsoft.Extensions.Logging;
using System.Text.RegularExpressions;

namespace Sql2Xls.Sql;

public class SqlStatement
{
    public string Statement { get; set; }

    public SqlStatement()
    {
        Statement = string.Empty;
    }
    public SqlStatement(string statement)
    {
        Statement = statement;
    }

    public static SqlStatement Load(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new InvalidOperationException("File name is missing");

        if (!File.Exists(filePath))
            throw new InvalidOperationException(string.Format("File {0} does not exist", filePath));

        using (FileStream fileStreamRead = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            using (StreamReader inputFile = new StreamReader(fileStreamRead))
            {
                string tmp = inputFile.ReadToEnd();
                SqlStatement statement = Parse(tmp);

                return statement;
            }
        }
    }

    public static SqlStatement Parse(string text)
    {
        text = RemoveComments(text, false);

        text = text.Replace("\r", " ");
        text = text.Replace("\n", " ");
        text = text.Replace("\t", " ");

        //remove repeated spaces
        while (text.Contains("  ")) text = text.Replace("  ", " ");

        if (IsLinkedServerQuery(text))
        {
            text = GetNestedQueryFromLinkedServerQuery(text);
            text = RemoveDoubleQuotes(text);
        }

        return new SqlStatement(text);
    }

    private static bool IsLinkedServerQuery(string text)
    {
        if (text.Contains("OPENQUERY("))
            return true;
        return false;
    }

    private static string GetNestedQueryFromLinkedServerQuery(string text)
    {
        string pattern =
            @"OPENQUERY[\S\s\w]*\(\[[\S\s\w]+\][ \s]*,[ \s\S]*'[ \s]*(SELECT[\s\w\S]*?)('\);)([\s\w\S]*OPENQUERY[\s\w\S]*|[\s\w\S]*)";

        MatchCollection matches = Regex.Matches(text, pattern,
            RegexOptions.IgnoreCase
            | RegexOptions.Multiline
            | RegexOptions.CultureInvariant);

        //Console.WriteLine("Matches found: {0}", matches.Count);

        if (matches.Count > 0)
            foreach (Match m in matches)
                return m.Groups[1].ToString();

        throw new InvalidOperationException("Nested SQL query extrtaction failed");
    }

    private static string RemoveDoubleQuotes(string sql)
    {
        return sql.Replace("''", "'");
    }

    //http://drizin.io/Removing-comments-from-SQL-scripts/
    static Regex everythingExceptNewLines = new Regex("[^\r\n]");
    public static string RemoveComments(string input, bool preservePositions, bool removeLiterals = false)
    {
        //based on http://stackoverflow.com/questions/3524317/regex-to-strip-line-comments-from-c-sharp/3524689#3524689
        var lineComments = @"--(.*?)\r?\n";
        var lineCommentsOnLastLine = @"--(.*?)$"; // because it's possible that there's no \r\n after the last line comment
                                                  // literals ('literals'), bracketedIdentifiers ([object]) and quotedIdentifiers ("object"), they follow the same structure:
                                                  // there's the start character, any consecutive pairs of closing characters are considered part of the literal/identifier, and then comes the closing character
        var literals = @"('(('')|[^'])*')"; // 'John', 'O''malley''s', etc
        var bracketedIdentifiers = @"\[((\]\])|[^\]])* \]"; // [object], [ % object]] ], etc
        var quotedIdentifiers = @"(\""((\""\"")|[^""])*\"")"; // "object", "object[]", etc - when QUOTED_IDENTIFIER is set to ON, they are identifiers, else they are literals
                                                              //var blockComments = @"/\*(.*?)\*/";  //the original code was for C#, but Microsoft SQL allows a nested block comments // //https://msdn.microsoft.com/en-us/library/ms178623.aspx

        //so we should use balancing groups // http://weblogs.asp.net/whaggard/377025
        var nestedBlockComments = @"/\*
                                 (?>
                                 /\*  (?<LEVEL>)      # On opening push level
                                 | 
                                 \*/ (?<-LEVEL>)     # On closing pop level
                                 |
                                 (?! /\* | \*/ ) . # Match any char unless the opening and closing strings   
                                 )+                         # /* or */ in the lookahead string
                                 (?(LEVEL)(?!))             # If level exists then fail
                                 \*/";

        string noComments = Regex.Replace(input,
            nestedBlockComments + "|" + lineComments + "|" + lineCommentsOnLastLine + "|" + literals + "|" + bracketedIdentifiers + "|" + quotedIdentifiers,
            me =>
            {
                if (me.Value.StartsWith("/*") && preservePositions)
                    return everythingExceptNewLines.Replace(me.Value, " "); // preserve positions and keep line-breaks // return new string(' ', me.Value.Length);
                else if (me.Value.StartsWith("/*") && !preservePositions)
                    return "";
                else if (me.Value.StartsWith("--") && preservePositions)
                    return everythingExceptNewLines.Replace(me.Value, " "); // preserve positions and keep line-breaks
                else if (me.Value.StartsWith("--") && !preservePositions)
                    return everythingExceptNewLines.Replace(me.Value, ""); // preserve only line-breaks // Environment.NewLine;
                else if (me.Value.StartsWith("[") || me.Value.StartsWith("\""))
                    return me.Value; // do not remove object identifiers ever
                else if (!removeLiterals) // Keep the literal strings
                    return me.Value;
                else if (removeLiterals && preservePositions) // remove literals, but preserving positions and line-breaks
                {
                    var literalWithLineBreaks = everythingExceptNewLines.Replace(me.Value, " ");
                    return "'" + literalWithLineBreaks.Substring(1, literalWithLineBreaks.Length - 2) + "'";
                }
                else if (removeLiterals && !preservePositions) // wrap completely all literals
                    return "''";
                else
                    throw new NotImplementedException();
            },
            RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace);
        return noComments;
    }
}
