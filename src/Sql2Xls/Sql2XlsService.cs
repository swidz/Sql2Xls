using Microsoft.Extensions.Logging;
using Sql2Xls.Excel;
using Sql2Xls.Sql;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;
using System.Security.AccessControl;
using System.Text.RegularExpressions;

namespace Sql2Xls;

public class Sql2XlsService : ISql2XlsService
{
    
    private readonly ISqlDataService _sqlService;
    private readonly ILogger<Sql2XlsService> _logger;
    public readonly ILoggerFactory _loggerFactory;
    public readonly ISql2XlsOptions _options;

    private ISql2XlsOptions Context 
    { 
        get { return _options; } 
    }

    private const int MAX_NAME_LENGTH = 30;
    private const string OUTPUT_FILE_EXTENSION = "xlsx";
    private const string SOURCE_SEARCH_PATTERN = "*.sql";

    private string sourceFolder = String.Empty;
    private string sourceSearchPattern = String.Empty;
    private string destinationFile = String.Empty;
    private string destinationFolder = String.Empty;

    private string[] files = null;

    public Sql2XlsService(ISql2XlsOptions options, ISqlDataService sqlService, ILoggerFactory loggerFactory)
    {
        _sqlService = sqlService ?? throw new ArgumentNullException(nameof(sqlService));
        _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
        _logger = _loggerFactory.CreateLogger<Sql2XlsService>();
    }

    private void CreateDocument(string datasetName, string sqlCommand, string outputFile)
    {
        DataTable dt = _sqlService.ExecuteQuery(
            Context.DatabaseProviderName, Context.ConnectionString,
            sqlCommand, Context.ConnectionTimeOut);

        dt.TableName = datasetName;
        PreprocessDataTable(dt);

        var excelContext = new ExcelExportContext()
        {
            SheetName = datasetName,
            ProviderName = Context.ExportEngine,
            FileName = outputFile,
            ODCConnectionString = Context.ConnectionString,
            ODCSqlStatement = sqlCommand,
            ODCTableName = datasetName
        };

        var factory = new ExcelExportFactory(_loggerFactory);
        var excelExport = factory.Create(excelContext);
        excelExport.LoadFromDataTable(dt);
    }

    private void CreateDocumentFromDataRecord(string datasetName, string sqlCommand, string outputFile)
    {
        var excelContext = new ExcelExportContext()
        {
            ProviderName = Context.ExportEngine,
            FileName = outputFile,
            ODCConnectionString = Context.ConnectionString,
            ODCSqlStatement = sqlCommand,
            ODCTableName = datasetName
        };

        var factory = new ExcelExportFactory(_loggerFactory);
        var excelExport = factory.Create(excelContext);

        _sqlService.ExecuteQuery(
            Context.DatabaseProviderName,
            Context.ConnectionString,
            sqlCommand,
            Context.ConnectionTimeOut,
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

    private void InitSource()
    {
        sourceFolder = String.IsNullOrEmpty(Context.Source)
            ? Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName)
            : Path.GetDirectoryName(Context.Source);

        if (String.IsNullOrEmpty(sourceFolder))
        {
            sourceFolder = Path.GetPathRoot(Context.Source);
        }
        _logger.LogTrace("Source folder is {0}", sourceFolder);

        sourceSearchPattern = SOURCE_SEARCH_PATTERN;
        if (!String.IsNullOrEmpty(Context.Source) && !String.IsNullOrEmpty(Path.GetFileName(Context.Source)))
        {
            sourceSearchPattern = Path.GetFileName(Context.Source);
        }

        _logger.LogTrace("Source pattern is {0}", sourceSearchPattern);

        if (!Directory.Exists(sourceFolder))
        {
            throw new InvalidOperationException(String.Format("Source directory {0} does not exist.", sourceFolder));
        }

    }

    private void InitDestination()
    {
        destinationFolder = String.IsNullOrEmpty(Context.Destination)
            ? sourceFolder
            : Path.GetDirectoryName(Context.Destination);

        if (String.IsNullOrEmpty(destinationFolder))
        {
            destinationFolder = Path.GetPathRoot(Context.Destination);
        }
        _logger.LogTrace("Output folder is {0}", destinationFolder);

        if (!String.IsNullOrEmpty(Context.Destination) && !String.IsNullOrEmpty(Path.GetFileName(Context.Destination)))
        {
            destinationFile = Path.GetFileName(Context.Destination);
        }

        if (!Context.CreateOutputFolder && !Directory.Exists(destinationFolder))
        {
            throw new InvalidOperationException(String.Format("Output directory {0} does not exist.", destinationFolder));
        }
    }

    private void InitZip()
    {
        if (!String.IsNullOrEmpty(Context.ZipOutputFolder))
        {
            if (!Context.CreateOutputFolder && !Directory.Exists(Context.ZipOutputFolder))
                throw new InvalidOperationException(String.Format("Zip output directory {0} does not exist.", Context.ZipOutputFolder));

            try
            {
                if (Context.CreateOutputFolder && !Directory.Exists(Context.ZipOutputFolder))
                {
                    _logger.LogTrace("Creating zip output directory {0}", Context.ZipOutputFolder);
                    Directory.CreateDirectory(Context.ZipOutputFolder);
                }
            }
            catch (Exception ex)
            {

                _logger.LogError("Error creating directory {0} {1}", Context.ZipOutputFolder, ex.Message);
                throw;
            }
        }
    }

    private void InitDb()
    {
        try
        {
            _sqlService.RegisterDbProviderFactory(Context.DatabaseProviderName);
        }
        catch (Exception ex)
        {
            _logger.LogError("Error registering database provider {0} {1}", Context.DatabaseProviderName, ex.Message);
            throw;
        }
    }

    private void InitFiles()
    {
        files = System.IO.Directory.GetFiles(sourceFolder, sourceSearchPattern);

        if (files.Length == 0)
        {
            _logger.LogWarning("No files matching {0} path", Context.Source);
        }

        if (files.Length > 1 && !String.IsNullOrEmpty(destinationFile))
        {
            destinationFile = String.Empty;
        }

        Array.Sort(files);
    }

    private void Init()
    {
        InitSource();
        InitDestination();
        InitZip();
        InitDb();
        InitFiles();
    }

    private void CreateDestinationFolders()
    {
        if (files.Length > 1 && !Directory.Exists(destinationFolder))
        {
            try
            {
                Directory.CreateDirectory(destinationFolder);
                _logger.LogTrace("Creating destination folder {0}", destinationFolder);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Creating destination folder {0}", destinationFolder);
                throw;
            }
        }
    }

    private string GetName(string file)
    {
        string name = Context.WorksheetName;
        if (String.IsNullOrWhiteSpace(Context.WorksheetName))
        {
            name = Path.GetFileNameWithoutExtension(file);
            name = name.Replace(" ", String.Empty);
            if (name.Length > MAX_NAME_LENGTH)
                name = name.Substring(0, MAX_NAME_LENGTH);
            name = name.Replace('-', '_');
        }

        return name;
    }

    private string GetDestinationFilePath(string file)
    {
        string outputFilename = Path.GetFileNameWithoutExtension(file);
        if (!String.IsNullOrEmpty(Context.OutputFileSuffix))
            outputFilename += Context.OutputFileSuffix;
        if (string.IsNullOrEmpty(Path.GetExtension(outputFilename)))
        {
            outputFilename += ".";
            outputFilename += OUTPUT_FILE_EXTENSION;
        }
        else
        {
            outputFilename = Path.ChangeExtension(outputFilename, OUTPUT_FILE_EXTENSION);
        }

        string destinationFilePath = String.IsNullOrEmpty(destinationFile)
            ? Path.Combine(destinationFolder, outputFilename)
            : Path.Combine(destinationFolder, destinationFile);

        return destinationFilePath;
    }

    private void CheckDestinationFilePath(string destinationFilePath)
    {
        if (!Context.Overwrite && File.Exists(destinationFilePath))
        {
            IOException ex = new IOException(String.Format("File {0} already exists, use -x switch to enable overwriting.", destinationFilePath));
            _logger.LogError(ex, "File {0} already exists", destinationFilePath);
            throw ex;
        }

        CheckWriteAccessToFolder(Path.GetDirectoryName(destinationFilePath));
    }

    private bool CheckWriteAccessToFolder(string folderPath)
    {
        try
        {
            DirectoryInfo di = new DirectoryInfo(folderPath);
            // Attempt to get a list of security permissions from the folder. 
            // This will raise an exception if the path is read only or do not have access to view the permissions. 

            //TODO https://stackoverflow.com/questions/49430088/check-access-permisions-in-c-sharp-on-linux
            DirectorySecurity ds = di.GetAccessControl();
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    public void Run()
    {
        Init();
        CreateDestinationFolders();

        bool hasError = false;
        var destinationFolders = new HashSet<string>(files.Length);
        var tasks = new List<Tuple<string, string, string>>(files.Length);

        foreach (var file in files)
        {
            _logger.LogInformation("Pre processing file {0}", file);

            string sourceFilePath = Path.Combine(sourceFolder, file);
            _logger.LogTrace("Source file path is {0}", sourceFilePath);

            string name = GetName(file);
            _logger.LogTrace("Dataset name is {0}", name);

            string destinationFilePath = GetDestinationFilePath(file);
            _logger.LogTrace("Output file is {0}", destinationFilePath);

            CheckDestinationFilePath(destinationFilePath);

            tasks.Add(Tuple.Create<string, string, string>(
                name,
                sourceFilePath,
                destinationFilePath));
        }

        if (Context.MaxDegreeOfParallelism == 1)
        {
            foreach (var task in tasks)
            {
                try
                {
                    _logger.LogInformation("Start processing file {0}", task.Item2);

                    var statement = SqlStatement.Load(task.Item2).Statement;
                    _logger.LogTrace("Statement is: {0}", statement);

                    CreateDocument(task.Item1, statement, task.Item3);

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
                    MaxDegreeOfParallelism = Context.MaxDegreeOfParallelism
                },
                (task) =>
                {
                    try
                    {
                        _logger.LogInformation("Start processing file {0}", task.Item2);
                        
                        var statement = SqlStatement.Load(task.Item2).Statement;
                        _logger.LogTrace("Statement is: {0}", statement);

                        CreateDocument(task.Item1, statement, task.Item3);
                        _logger.LogInformation("Output file {0} was created", task.Item3);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error processing file {0}", task.Item2);
                        hasError = true;
                    }
                });
        }

        if (!String.IsNullOrEmpty(Context.ZipOutputFolder))
        {
            if (!Context.CreateOutputFolder && !Directory.Exists(Context.ZipOutputFolder))
                throw new InvalidOperationException(String.Format("Zip output directory {0} does not exist.", Context.ZipOutputFolder));

            try
            {
                if (Context.CreateOutputFolder && !Directory.Exists(Context.ZipOutputFolder))
                {
                    _logger.LogTrace("Creating zip output directory {0}", Context.ZipOutputFolder);
                    Directory.CreateDirectory(Context.ZipOutputFolder);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating directory {0}", Context.ZipOutputFolder);
                throw;
            }

            foreach (string folder in destinationFolders)
            {
                string zipOutputPath = Path.Combine(Context.ZipOutputFolder, Path.ChangeExtension(folder, "zip"));
                this.CreateZipFile(folder, zipOutputPath);
            }
        }

        if (hasError)
        {
            _logger.LogError("Process completed with errors. Please check the log file {0}", Path.Combine(destinationFolder, Context.LogFileName));
        }
        else
        {
            _logger.LogInformation("Process completed.");
        }
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

}