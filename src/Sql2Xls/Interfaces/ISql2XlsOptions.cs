namespace Sql2Xls.Interfaces
{
    public interface ISql2XlsOptions
    {
        string ConnectionString { get; }
        int ConnectionTimeOut { get; }
        bool CreateOutputFolder { get; }
        string DatabaseProviderName { get; }
        string Destination { get; }
        string ExportEngine { get; }
        string LogFileName { get; }
        string LogFullPath { get; }
        int LogLevel { get; }
        int MaxDegreeOfParallelism { get; }
        string OutputFileSuffix { get; }
        bool Overwrite { get; }
        string Source { get; }
        string WorksheetName { get; }
        string ZipOutputFolder { get; }
    }
}