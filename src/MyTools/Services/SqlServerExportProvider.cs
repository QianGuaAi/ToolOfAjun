using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public sealed class SqlServerExportProvider : ISqlExportProvider
    {
        public SqlProviderKind ProviderKind => SqlProviderKind.SqlServer;

        public Task TestConnectionAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
            => SqlExportService.TestConnectionAsync(options, cancellationToken);

        public Task<List<DatabaseItem>> GetDatabasesAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
            => SqlExportService.GetDatabasesAsync(options, cancellationToken);

        public Task<List<TableItem>> GetTablesAsync(SqlServerConnectionOptions options, string databaseName, CancellationToken cancellationToken)
            => SqlExportService.GetTablesAsync(options, databaseName, cancellationToken);

        public Task<ExportResult> ExportTableAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            TableItem table,
            string filePath,
            CancellationToken cancellationToken)
            => SqlExportService.ExportTableAsync(options, databaseName, table, filePath, cancellationToken);

        public Task<DataTable> ExecuteQueryAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            string sql,
            CancellationToken cancellationToken)
            => SqlExportService.ExecuteQueryAsync(options, databaseName, sql, cancellationToken);
    }
}
