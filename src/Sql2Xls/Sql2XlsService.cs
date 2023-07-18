using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.EMMA;
using Microsoft.Extensions.Logging;
using Sql2Xls.Excel;
using Sql2Xls.Sql;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text.RegularExpressions;

namespace Sql2Xls;

public class Sql2XlsService : ISql2XlsService
{
    private readonly ISqlDataService _sqlService;
    private readonly ILogger<Sql2XlsService> _logger;
    private readonly ILoggerFactory _loggerFactory;

    private const int MAX_NAME_LENGTH = 30;
    private const string OUTPUT_FILE_EXTENSION = "xlsx";
    private const string SOURCE_SEARCH_PATTERN = "*.sql";

    public Sql2XlsService(ISqlDataService sqlService, ILoggerFactory loggerFactory)
    {
        _sqlService = sqlService ?? throw new ArgumentNullException(nameof(sqlService));
        _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
        _logger = _loggerFactory.CreateLogger<Sql2XlsService>();
    }

    public void Run(ISql2XlsOptions options)
    {

        var parms = new Sql2XlsServiceParameters(options);

        Init(parms);
        CreateDestinationFolders(parms);

        bool hasError = false;
        var destinationFolders = new HashSet<string>(parms.Files.Length);
        var tasks = new List<Tuple<string, string, string>>(parms.Files.Length);

        foreach (var file in parms.Files)
        {
            _logger.LogInformation("Pre processing file {0}", file);

            var sourceFilePath = Path.Combine(parms.SourceFolder, file);
            _logger.LogTrace("Source file path is {0}", sourceFilePath);

            var name = GetName(file, parms);
            _logger.LogTrace("Dataset name is {0}", name);

            var destinationFilePath = GetDestinationFilePath(file, parms);
            _logger.LogTrace("Output file is {0}", destinationFilePath);

            CheckOverwrite(destinationFilePath, parms);

            tasks.Add(Tuple.Create<string, string, string>(
                name,
                sourceFilePath,
                destinationFilePath));
        }

        if (parms.Options.MaxDegreeOfParallelism == 1)
        {
            foreach (var task in tasks)
            {
                try
                {
                    _logger.LogInformation("Start processing file {0}", task.Item2);

                    var statement = SqlStatement.Load(task.Item2).Statement;
                    _logger.LogTrace("Statement is: {0}", statement);

                    CreateDocument(task.Item1, statement, task.Item3, parms);

                    _logger.LogInformation("Output file {0} was created", task.Item3);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing file {0}", task.Item2);
                    hasError = true;
                }
            }
        }
        else
        {
            Parallel.ForEach(tasks,
                new ParallelOptions
                {
                    MaxDegreeOfParallelism = parms.Options.MaxDegreeOfParallelism
                },
                (task) =>
                {
                    try
                    {
                        _logger.LogInformation("Start processing file {0}", task.Item2);

                        var statement = SqlStatement.Load(task.Item2).Statement;
                        _logger.LogTrace("Statement is: {0}", statement);

                        CreateDocument(task.Item1, statement, task.Item3, parms);
                        _logger.LogInformation("Output file {0} was created", task.Item3);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error processing file {0}", task.Item2);
                        hasError = true;
                    }
                });
        }

        if (!String.IsNullOrEmpty(parms.Options.ZipOutputFolder))
        {
            if (!parms.Options.CreateOutputFolder && !Directory.Exists(parms.Options.ZipOutputFolder))
                throw new InvalidOperationException(String.Format("Zip output directory {0} does not exist.", parms.Options.ZipOutputFolder));

            try
            {
                if (parms.Options.CreateOutputFolder && !Directory.Exists(parms.Options.ZipOutputFolder))
                {
                    _logger.LogTrace("Creating zip output directory {0}", parms.Options.ZipOutputFolder);
                    Directory.CreateDirectory(parms.Options.ZipOutputFolder);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating directory {0}", parms.Options.ZipOutputFolder);
                throw;
            }

            foreach (string folder in destinationFolders)
            {
                string zipOutputPath = Path.Combine(parms.Options.ZipOutputFolder, Path.ChangeExtension(folder, "zip"));
                this.CreateZipFile(folder, zipOutputPath);
            }
        }

        if (hasError)
        {
            _logger.LogError("Process completed with errors. Please check the log file {0}", Path.Combine(parms.DestinationFolder, parms.Options.LogFileName));
        }
        else
        {
            _logger.LogInformation("Process completed.");
        }
    }

    private void CreateDocument(string datasetName, string sqlCommand, string outputFile, Sql2XlsServiceParameters parms)
    {
        DataTable dt = _sqlService.ExecuteQuery(
            parms.Options.DatabaseProviderName, parms.Options.ConnectionString,
            sqlCommand, parms.Options.ConnectionTimeOut);

        dt.TableName = datasetName;
        PreprocessDataTable(dt);

        var excelContext = new ExcelExportContext()
        {
            SheetName = datasetName,
            ProviderName = parms.Options.ExportEngine,
            FileName = outputFile,
            ODCConnectionString = parms.Options.ConnectionString,
            ODCSqlStatement = sqlCommand,
            ODCTableName = datasetName
        };

        var factory = new ExcelExportFactory(_loggerFactory);
        var excelExport = factory.Create(excelContext);

        excelExport.LoadFromDataTable(dt);
    }

    private void CreateDocumentFromDataRecord(string datasetName, string sqlCommand, string outputFile, Sql2XlsServiceParameters parms)
    {
        var excelContext = new ExcelExportContext()
        {
            ProviderName = parms.Options.ExportEngine,
            FileName = outputFile,
            ODCConnectionString = parms.Options.ConnectionString,
            ODCSqlStatement = sqlCommand,
            ODCTableName = datasetName
        };

        var factory = new ExcelExportFactory(_loggerFactory);
        var excelExport = factory.Create(excelContext);

        _sqlService.ExecuteQuery(
            parms.Options.DatabaseProviderName,
            parms.Options.ConnectionString,
            sqlCommand,
            parms.Options.ConnectionTimeOut,
            (dataRecord) =>
            {
                excelExport.AddDataRecord(dataRecord);
            });

        excelExport.Close();
    }

    private void PreprocessDataTable(DataTable dataTable)
    {
        char[] newLineSeparators = new char[] { '\r', '\n' };
        Regex regex = new Regex(@"\p{C}+", RegexOptions.Compiled);

        int numberOfCols = dataTable.Columns.Count;
        foreach (System.Data.DataRow dsrow in dataTable.Rows)
        {
            for (int colIndex = 0; colIndex < numberOfCols; colIndex++)
            {
                DataColumn column = dataTable.Columns[colIndex];
                if (column.DataType == typeof(String))
                {
                    //remove control characters
                    //replace new line with environment new line
                    //remove repeated new line chars
                    string resultValue = String.Join(
                            Environment.NewLine, dsrow[colIndex].ToString()
                                .Split(newLineSeparators, StringSplitOptions.RemoveEmptyEntries)
                                .Select(line => regex.Replace(line, String.Empty)));

                    dsrow[colIndex] = resultValue;
                }
            }
        }
    }

    private void InitSource(Sql2XlsServiceParameters parms)
    {
        var sourceFolder = String.IsNullOrEmpty(parms.Options.Source)
            ? Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
            : Path.GetDirectoryName(parms.Options.Source);

        if (String.IsNullOrEmpty(sourceFolder))
        {
            sourceFolder = Path.GetPathRoot(parms.Options.Source);
        }
        _logger.LogTrace("Source folder is {0}", sourceFolder);
        parms.SourceFolder = sourceFolder;

        var sourceSearchPattern = SOURCE_SEARCH_PATTERN;
        if (!String.IsNullOrEmpty(parms.Options.Source) && !String.IsNullOrEmpty(Path.GetFileName(parms.Options.Source)))
        {
            sourceSearchPattern = Path.GetFileName(parms.Options.Source);
        }
        _logger.LogTrace("Source pattern is {0}", sourceSearchPattern);
        parms.SourceSearchPattern = sourceSearchPattern;

        if (!Directory.Exists(sourceFolder))
        {
            throw new InvalidOperationException(String.Format("Source directory {0} does not exist.", sourceFolder));
        }

    }

    private void InitDestination(Sql2XlsServiceParameters parms)
    {
        var destinationFolder = String.IsNullOrEmpty(parms.Options.Destination)
            ? parms.SourceFolder
            : Path.GetDirectoryName(parms.Options.Destination);

        if (String.IsNullOrEmpty(destinationFolder))
        {
            destinationFolder = Path.GetPathRoot(parms.Options.Destination);
        }
        _logger.LogTrace("Output folder is {0}", destinationFolder);
        parms.DestinationFolder = destinationFolder;

        if (!String.IsNullOrEmpty(parms.Options.Destination) && !String.IsNullOrEmpty(Path.GetFileName(parms.Options.Destination)))
        {
            parms.DestinationFile = Path.GetFileName(parms.Options.Destination);
        }

        if (!parms.Options.CreateOutputFolder && !Directory.Exists(destinationFolder))
        {
            throw new InvalidOperationException(String.Format("Output directory {0} does not exist.", destinationFolder));
        }
    }

    private void InitZip(Sql2XlsServiceParameters parms)
    {
        if (!String.IsNullOrEmpty(parms.Options.ZipOutputFolder))
        {
            if (!parms.Options.CreateOutputFolder && !Directory.Exists(parms.Options.ZipOutputFolder))
                throw new InvalidOperationException(String.Format("Zip output directory {0} does not exist.", parms.Options.ZipOutputFolder));

            try
            {
                if (parms.Options.CreateOutputFolder && !Directory.Exists(parms.Options.ZipOutputFolder))
                {
                    _logger.LogTrace("Creating zip output directory {0}", parms.Options.ZipOutputFolder);
                    Directory.CreateDirectory(parms.Options.ZipOutputFolder);
                }
            }
            catch (Exception ex)
            {

                _logger.LogError("Error creating directory {0} {1}", parms.Options.ZipOutputFolder, ex.Message);
                throw;
            }
        }
    }

    private void InitDb(Sql2XlsServiceParameters parms)
    {
        try
        {
            _sqlService.RegisterDbProviderFactory(parms.Options.DatabaseProviderName);
        }
        catch (Exception ex)
        {
            _logger.LogError("Error registering database provider {0} {1}", parms.Options.DatabaseProviderName, ex.Message);
            throw;
        }
    }

    private void InitFiles(Sql2XlsServiceParameters parms)
    {
        parms.Files = System.IO.Directory.GetFiles(parms.SourceFolder, parms.SourceSearchPattern);

        if (parms.Files.Length == 0)
        {
            _logger.LogWarning("No files matching {0} path", parms.Options.Source);
        }

        if (parms.Files.Length > 1 && !String.IsNullOrEmpty(parms.DestinationFile))
        {
            parms.DestinationFile = String.Empty;
        }

        Array.Sort(parms.Files);
    }

    private void Init(Sql2XlsServiceParameters parms)
    {
        InitSource(parms);
        InitDestination(parms);
        InitZip(parms);
        InitDb(parms);
        InitFiles(parms);
    }

    private void CreateDestinationFolders(Sql2XlsServiceParameters parms)
    {
        if (parms.Files.Length > 1 && !Directory.Exists(parms.DestinationFolder))
        {
            try
            {
                Directory.CreateDirectory(parms.DestinationFolder);
                _logger.LogTrace("Creating destination folder {0}", parms.DestinationFolder);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Creating destination folder {0}", parms.DestinationFolder);
                throw;
            }
        }
    }

    private string GetName(string file, Sql2XlsServiceParameters parms)
    {
        string name = parms.Options.WorksheetName;
        if (String.IsNullOrWhiteSpace(parms.Options.WorksheetName))
        {
            name = Path.GetFileNameWithoutExtension(file);
            name = name.Replace(" ", String.Empty);
            if (name.Length > MAX_NAME_LENGTH)
                name = name.Substring(0, MAX_NAME_LENGTH);
            name = name.Replace('-', '_');
        }

        return name;
    }

    private string GetDestinationFilePath(string file, Sql2XlsServiceParameters parms)
    {
        string outputFilename = Path.GetFileNameWithoutExtension(file);
        if (!String.IsNullOrEmpty(parms.Options.OutputFileSuffix))
            outputFilename += parms.Options.OutputFileSuffix;
        if (string.IsNullOrEmpty(Path.GetExtension(outputFilename)))
        {
            outputFilename += ".";
            outputFilename += OUTPUT_FILE_EXTENSION;
        }
        else
        {
            outputFilename = Path.ChangeExtension(outputFilename, OUTPUT_FILE_EXTENSION);
        }

        string destinationFilePath = String.IsNullOrEmpty(parms.DestinationFile)
            ? Path.Combine(parms.DestinationFolder, outputFilename)
            : Path.Combine(parms.DestinationFolder, parms.DestinationFile);

        return destinationFilePath;
    }

    private void CheckOverwrite(string destinationFilePath, Sql2XlsServiceParameters parms)
    {
        if (!parms.Options.Overwrite && File.Exists(destinationFilePath))
        {
            var ex = new IOException(String.Format("File {0} already exists, use -x switch to enable overwriting.", destinationFilePath));
            _logger.LogError(ex, "File {0} already exists", destinationFilePath);
            throw ex;
        }

        CheckWriteAccessToFolder(Path.GetDirectoryName(destinationFilePath));
    }

    private bool CheckWriteAccessToFolder(string folderPath)
    {

        var di = new DirectoryInfo(folderPath);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            try
            {
                // Attempt to get a list of security permissions from the folder. 
                // This will raise an exception if the path is read only or do not have access to view the permissions. 
        
                DirectorySecurity ds = di.GetAccessControl();
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }
        

        var mode = di.UnixFileMode;
        if (mode.HasFlag(UnixFileMode.UserRead) && mode.HasFlag(UnixFileMode.UserWrite))
        {
            return true;
        }

        return false;

        /*
        try
        {
            File.SetUnixFileMode(folderPath, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        */

        
    }

    private void CreateZipFile(string sourceDirectoryName, string destinationArchiveFileName)
    {
        try
        {
            if (File.Exists(destinationArchiveFileName))
            {
                _logger.LogTrace("Deleting existing zip file {0}", destinationArchiveFileName);
                File.Delete(destinationArchiveFileName);
            }

            _logger.LogTrace("Creating zip file {0}", destinationArchiveFileName);
            ZipFile.CreateFromDirectory(sourceDirectoryName, destinationArchiveFileName, CompressionLevel.Fastest, false);

        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error creating zip file for directory {0}", sourceDirectoryName);
            throw;
        }
    }

    internal class Sql2XlsServiceParameters
    {
        private readonly ISql2XlsOptions _options;

        public Sql2XlsServiceParameters(ISql2XlsOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
        }

        public string SourceFolder { get; set; } = String.Empty;
        public string SourceSearchPattern { get; set; } = String.Empty;
        public string DestinationFile { get; set; } = String.Empty;
        public string DestinationFolder { get; set; } = String.Empty;
        public string[] Files { get; set; } = null;
        public ISql2XlsOptions Options { get { return _options; } }


    }
}