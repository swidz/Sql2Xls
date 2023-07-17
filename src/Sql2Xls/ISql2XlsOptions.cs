namespace Sql2Xls
{
    public interface ISql2XlsOptions
    {
        string ConnectionString { get; set; }
        int ConnectionTimeOut { get; set; }
        bool CreateOutputFolder { get; set; }
        string DatabaseProviderName { get; set; }
        string Destination { get; set; }
        string ExportEngine { get; set; }
        string LogFileName { get; set; }
        string LogFullPath { get; }
        int LogLevel { get; set; }
        int MaxDegreeOfParallelism { get; set; }
        string OutputFileSuffix { get; set; }
        bool Overwrite { get; set; }
        string Source { get; set; }
        string WorksheetName { get; set; }
        string ZipOutputFolder { get; set; }
    }
}