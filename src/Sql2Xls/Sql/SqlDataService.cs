using Microsoft.Extensions.Logging;
using Npgsql;
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.SqlClient;

namespace Sql2Xls.Sql;

public class SqlDataService : ISqlDataService
{
    private readonly ILogger<SqlDataService> _logger;

    public SqlDataService(ILogger<SqlDataService> logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Given a provider name and connection string, create the <c>DbProviderFactory</c> and <c>DbConnection</c>
    /// </summary>
    /// <param name="providerName">Provider name</param>
    /// <param name="connectionString">Connection string</param>
    /// <returns>Returns a <c>DbConnection</c> on success; null on failure</returns>
    public DbConnection CreateDbConnection(string providerName, string connectionString)
    {
        DbConnection connection = null;
        if (connectionString != null)
        {
            try
            {
                DbProviderFactory factory = DbProviderFactories.GetFactory(providerName);
                connection = factory.CreateConnection();
                connection.ConnectionString = connectionString;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occured when creating a database connection using {0} provider name and {1} as connection string", providerName, connectionString);
                throw;
            }
        }

        return connection;
    }

    public static void RegisterDbProviderFactories()
    {
        DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.MSSQL_NAME, SqlClientFactory.Instance);
        DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.POSTEGRESQL_NAME, NpgsqlFactory.Instance);
        DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.ODBC_NAME, OdbcFactory.Instance);
    }

    public void RegisterDbProviderFactory(string providerName)
    {
        switch (providerName)
        {
            case DbProviderFactoryHelper.MSSQL_NAME:
                DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.MSSQL_NAME, System.Data.SqlClient.SqlClientFactory.Instance);
                break;

            case DbProviderFactoryHelper.POSTEGRESQL_NAME:
                DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.POSTEGRESQL_NAME, NpgsqlFactory.Instance);
                break;

            case DbProviderFactoryHelper.ODBC_NAME:
                DbProviderFactories.RegisterFactory(DbProviderFactoryHelper.ODBC_NAME, System.Data.Odbc.OdbcFactory.Instance);
                break;

            default:
                throw new NotImplementedException($"Provider name '{providerName}' is not supported");
        }
    }

    public DataTable ExecuteQuery(string providerName, string connectionString, string statement, int timeout)
    {
        DataTable result = null;

        DbProviderFactory factory = DbProviderFactories.GetFactory(providerName);

        using (DbConnection conn = CreateDbConnection(providerName, connectionString))
        {
            try
            {
                conn.Open();

                using (DbCommand command = conn.CreateCommand())
                {
                    command.CommandText = statement;
                    command.CommandType = CommandType.Text;
                    command.CommandTimeout = timeout;

                    result = new DataTable();

                    using (DbDataAdapter da = factory.CreateDataAdapter())
                    {
                        da.SelectCommand = command;
                        da.Fill(result);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occured running statement {0}", statement);
                throw;
            }
            finally
            {
                conn.Close();
            }
        }
        return result;
    }

    public void ExecuteQuery(string providerName, string connectionString, string statement, int timeout, Action<IDataRecord> action)
    {
        using (DbConnection conn = CreateDbConnection(providerName, connectionString))
        {
            try
            {
                conn.Open();
                using (DbCommand command = conn.CreateCommand())
                {
                    command.CommandText = statement;
                    command.CommandType = CommandType.Text;
                    command.CommandTimeout = timeout;

                    var dbReader = command.ExecuteReader();
                    while (dbReader.Read())
                    {
                        action.Invoke(dbReader);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occured running statement {0}", statement);
                throw;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
