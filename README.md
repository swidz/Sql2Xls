# Sql2Xls
Export SQL query results to Microsoft Excel

1. Create MS Excel files (.xlsx) based on SQL queries stored in specified folder
2. Excel files are compatible with MS Office 2007, 2010 and 2013 and Microsoft 365
3. Excel files are compatible with (can be read by) Ole.Db (Provider=Microsoft.ACE.OLEDB.12.0)
4. Columns are formatted based on data types returned from sql query (numbers, integers, text, datetime)
5. String values and dates are saved as shared strings in Excel
6. Output files are Mac compatible
7. SQL queries can be stored in .sql files, the tool runs against a folder path containing multiple .sql files
8. Input files can contain comments (both block and single line comments)
9. SQL queries can be run against the following databases: Microsoft SQL Server, PostgreSQL, ODBC
10. Logging (various levels, console and log file with a path specified by user) - the library is using LibLog and SeriLog with Console and File sinks
11. Queries can be executed in Parallel
12. Excel generation using: Performance adapter OpenXmlWriter (SAX) or standard OpenXml (DOM/ODC) (slower but more stable)
