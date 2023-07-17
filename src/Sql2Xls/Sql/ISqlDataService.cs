using System.Data;
using System.Data.Common;

namespace Sql2Xls.Sql;

public interface ISqlDataService
{
    DbConnection CreateDbConnection(string providerName, string connectionString);
    DataTable ExecuteQuery(string providerName, string connectionString, string statement, int timeout);
    void ExecuteQuery(string providerName, string connectionString, string statement, int timeout, Action<IDataRecord> action);
    void RegisterDbProviderFactory(string providerName);
}
