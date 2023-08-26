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
