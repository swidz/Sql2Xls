using CommandLine;
using Sql2Xls.Interfaces;

namespace Sql2Xls;

public class Sql2XlsOptions : ISql2XlsOptions
{
    [Option('c', "connection",
        Required = true,
        HelpText = "Connection string")]
    public string ConnectionString { get; set; }

    [Option('e', "engine",
        Required = false,
        Default = "SAX",
        HelpText = "Excel export provider [\"SAX\" | \"ODC\" | \"LEGACY\"]")]
    public string ExportEngine { get; set; }

    [Option('f', "createfolder",
        Required = false,
        Default = false,
        HelpText = @"If set, the output folder is created if it does not exist")]
    public bool CreateOutputFolder { get; set; }

    [Option('i', "input",
        Required = true,
        //Default = "*.sql",
        HelpText = @"Source file path, including filename wildcards e.g. C:\Folder\*.sql ")]
    public string Source { get; set; }

    [Option('l', "logfile",
        Required = false,
        Default = "SQL2XML.log",
        HelpText = "Log file name which is created in the output folder")]
    public string LogFileName { get; set; }

    [Option('m', "maxdop",
        Required = false,
        Default = 1,
        HelpText = "Sets number of threads for parallel query processing.")]
    public int MaxDegreeOfParallelism { get; set; }

    [Option('n', "name",
        Required = false,
        Default = "",
        HelpText = "Sets dataset and worksheet name. If not set, the name is derived from the source file name.")]
    public string WorksheetName { get; set; }

    [Option('o', "output",
        Required = true,
    //Default = "",
        HelpText = @"Output path. It can be either a folder eg. C:\OutputFolder\ or a output file (the second case works only for a pattern resulting in a single input file)")]
    public string Destination { get; set; }

    [Option('p', "provider",
        Required = false,
        Default = "System.Data.SqlClient",
        HelpText = "Database provider name [\"System.Data.SqlClient\" | \"Npgsql\" | \"System.Data.ODBC\"]")]
    public string DatabaseProviderName { get; set; }

    [Option('s', "suffix",
        Required = false,
        Default = "",
        HelpText = "Adds given suffix to output filenames")]
    public string OutputFileSuffix { get; set; }

    [Option('t', "timeout",
        Required = false,
        Default = 1200,
        HelpText = "Connection timeout")]
    public int ConnectionTimeOut { get; set; }

    [Option('v', "loglevel",
        Required = false,
        Default = 2,
        HelpText = "Log level: 0=Trace 1=Debug 2=Info 3=Warning 4=Error 5=Critical")]
    public int LogLevel { get; set; }

    [Option('x', "overwrite",
        Required = false,
        Default = false,
        HelpText = "Overwrite existing files")]
    public bool Overwrite { get; set; }

    [Option('z', "zip",
        Required = false,
        Default = "",
        HelpText = "If set, files in the destination folder are zipped, and the archive file is placed in given location")]
    public string ZipOutputFolder { get; set; }

    public string LogFullPath
    {
        get { return Path.Combine(Path.GetDirectoryName(Destination), LogFileName); }
    }

    public Sql2XlsOptions()
    {
        DatabaseProviderName = "System.Data.SqlClient";
        ExportEngine = "SAX";
        LogLevel = 2;
        Source = "*.sql";
        LogFileName = "SQL2XLS.log";
        MaxDegreeOfParallelism = 1;
        ConnectionTimeOut = 1200;
        WorksheetName = String.Empty;
    }
}
