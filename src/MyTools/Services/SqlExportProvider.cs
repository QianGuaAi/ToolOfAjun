using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public enum SqlProviderKind
    {
        SqlServer = 0,
        PostgreSql = 1,
        MySql = 2
    }

    public interface ISqlExportProvider
    {
        SqlProviderKind ProviderKind { get; }

        Task TestConnectionAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken);

        Task<List<DatabaseItem>> GetDatabasesAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken);

        Task<List<TableItem>> GetTablesAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            CancellationToken cancellationToken);

        Task<ExportResult> ExportTableAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            TableItem table,
            string filePath,
            CancellationToken cancellationToken);

        Task<DataTable> ExecuteQueryAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            string sql,
            CancellationToken cancellationToken);
    }

    public static class SqlExportProviderFactory
    {
        private static readonly IReadOnlyDictionary<SqlProviderKind, ISqlExportProvider> Providers =
            new Dictionary<SqlProviderKind, ISqlExportProvider>
            {
                [SqlProviderKind.SqlServer] = new SqlServerExportProvider(),
                [SqlProviderKind.PostgreSql] = new PostgreSqlExportProvider(),
                [SqlProviderKind.MySql] = new MySqlExportProvider()
            };

        public static ISqlExportProvider GetProvider(SqlProviderKind providerKind)
        {
            if (Providers.TryGetValue(providerKind, out var provider))
            {
                return provider;
            }

            throw new NotSupportedException($"不支持的数据库类型：{providerKind}");
        }
    }
}
