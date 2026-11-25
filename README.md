# Sql2Xls
Export SQL query results to Microsoft Excel

1. Create MS Excel files (.xlsx) based on SQL queries stored in text files in specified folder
2. Excel files are compatible with MS Office 2007, 2010 and 2013 and Microsoft 365
3. Tabular data shape - files can be read by Ole.Db (e.g. Provider=Microsoft.ACE.OLEDB.12.0)
4. Columns are formatted based on data types returned from sql query (numbers, integers, text, datetime)
5. String values and dates are saved as shared strings in Excel (limits file size && speeds up opening large files)
6. Output files are Mac compatible
7. SQL queries can be stored in .sql files, the tool runs against a folder path containing multiple .sql files
8. Input files can contain comments (both block and single line comments)
9. SQL queries can be run against the following databases: Microsoft SQL Server, PostgreSQL, ODBC
10. Logging (various levels, console and log file with a path specified by user) - the library is using LibLog and SeriLog with Console and File sinks
11. Queries can be executed in Parallel (use Maxdop parameter)
12. Excel generation using: Fast streaming adapter using OpenXmlWriter (SAX) or standard OpenXml (DOM/ODC) (slower)


Sql2Xls.exe
Sql2Xls 2.0.0.0+de3020b3b106b2664d70584c723a825041a7f16a
2023-2024 (c) Infovision Sebastian Widz

  -c, --connection      Required. Connection string
  -e, --engine          (Default: SAX) Excel export provider ["SAX" | "ODC" |
                        "LEGACY" | "LARGEXLSX"]
  -f, --createfolder    (Default: false) If set, the output folder is created if
                        it does not exist
  -i, --input           Required. Source file path, including filename wildcards
                        e.g. C:\Folder\*.sql
  -l, --logfile         (Default: SQL2XML.log) Log file name which is created in
                        the output folder
  -m, --maxdop          (Default: 1) Sets number of threads for parallel query
                        processing.
  -n, --name            (Default: ) Sets dataset and worksheet name. If not set,
                        the name is derived from the source file name.
  -o, --output          Required. Output path. It can be either a folder eg.
                        C:\OutputFolder\ or a output file (the second case works
                        only for a pattern resulting in a single input file)
  -p, --provider        (Default: System.Data.SqlClient) Database provider name
                        ["System.Data.SqlClient" | "Npgsql" |
                        "System.Data.ODBC"]
  -s, --suffix          (Default: ) Adds given suffix to output filenames
  -t, --timeout         (Default: 1200) Connection timeout
  -v, --loglevel        (Default: 2) Log level: 0=Trace 1=Debug 2=Info 3=Warning
                        4=Error 5=Critical
  -w, --password        Password protect output files
  -x, --overwrite       (Default: false) Overwrite existing files
  -z, --zip             (Default: ) If set, files in the destination folder are
                        zipped, and the archive file is placed in given location
  --help                Display this help screen.
  --version             Display version information.
