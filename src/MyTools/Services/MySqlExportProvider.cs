using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using MySqlConnector;

namespace MyTools.Services
{
    public sealed class MySqlExportProvider : ISqlExportProvider
    {
        public SqlProviderKind ProviderKind => SqlProviderKind.MySql;

        public async Task TestConnectionAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            AppLogService.Information("Testing MySQL connection to {ServerAddress}:{Port}", options.ServerAddress, NormalizePort(options.Port, 3306));
            using (var connection = new MySqlConnection(BuildMasterConnectionString(options)))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
            }
        }

        public async Task<List<DatabaseItem>> GetDatabasesAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            var result = new List<DatabaseItem>();
            using (var connection = new MySqlConnection(BuildMasterConnectionString(options)))
            using (var command = new MySqlCommand("SHOW DATABASES;", connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        var name = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            result.Add(new DatabaseItem { Name = name });
                        }
                    }
                }
            }

            AppLogService.Information("Loaded {DatabaseCount} MySQL databases from {ServerAddress}", result.Count, options.ServerAddress);
            return result;
        }

        public async Task<List<TableItem>> GetTablesAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);
            var result = new List<TableItem>();
            using (var connection = new MySqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new MySqlCommand("SHOW TABLES;", connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        var tableName = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(tableName))
                        {
                            result.Add(new TableItem
                            {
                                SchemaName = databaseName,
                                TableName = tableName
                            });
                        }
                    }
                }
            }

            AppLogService.Information("Loaded {TableCount} MySQL tables from {DatabaseName}", result.Count, databaseName);
            return result;
        }

        public async Task<ExportResult> ExportTableAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            TableItem table,
            string filePath,
            CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);
            ValidateTable(table);
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("导出文件路径不能为空。", nameof(filePath));
            }

            var rowCount = await GetRowCountAsync(options, databaseName, table, cancellationToken).ConfigureAwait(false);
            if (rowCount > SqlExportService.ExcelWorksheetRowLimit)
            {
                throw new InvalidOperationException("该表数据量超过 Excel 单工作表上限，当前版本暂不支持自动分片导出，请改为筛选后导出或后续扩展 CSV/多 Sheet 功能。");
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath) ?? AppDomain.CurrentDomain.BaseDirectory);
            var selectSql = $"SELECT * FROM {GetQualifiedTableName(table)};";
            var dataTable = new DataTable(table.DisplayName);
            using (var connection = new MySqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new MySqlCommand(selectSql, connection))
            {
                command.CommandTimeout = 0;
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    dataTable.Load(reader);
                }
            }

            AppLogService.Information(
                "MySQL export completed for {DatabaseName}.{SchemaName}.{TableName} with {RowCount} rows",
                databaseName,
                table.SchemaName,
                table.TableName,
                rowCount);
            return await SqlExportService.ExportDataTableAsync(dataTable, table.DisplayName, filePath, cancellationToken).ConfigureAwait(false);
        }

        public async Task<DataTable> ExecuteQueryAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            string sql,
            CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);
            if (string.IsNullOrWhiteSpace(sql))
            {
                throw new InvalidOperationException("请输入 SQL 语句。");
            }

            var dataTable = new DataTable();
            using (var connection = new MySqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new MySqlCommand(sql, connection))
            {
                command.CommandTimeout = 120;
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                {
                    dataTable.Load(reader);
                }
            }

            return dataTable;
        }

        private async Task<long> GetRowCountAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            TableItem table,
            CancellationToken cancellationToken)
        {
            var sql = $"SELECT COUNT(1) FROM {GetQualifiedTableName(table)};";
            using (var connection = new MySqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new MySqlCommand(sql, connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                var result = await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);
                return Convert.ToInt64(result, CultureInfo.InvariantCulture);
            }
        }

        private static string GetQualifiedTableName(TableItem table)
        {
            return $"{EscapeIdentifier(table.SchemaName)}.{EscapeIdentifier(table.TableName)}";
        }

        private static string EscapeIdentifier(string value)
        {
            return "`" + (value ?? string.Empty).Replace("`", "``") + "`";
        }

        private static void ValidateConnectionOptions(SqlServerConnectionOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.ServerAddress))
            {
                throw new InvalidOperationException("请输入服务器地址。");
            }

            if (string.IsNullOrWhiteSpace(options.Username))
            {
                throw new InvalidOperationException("请输入用户名。");
            }
        }

        private static void ValidateDatabaseName(string databaseName)
        {
            if (string.IsNullOrWhiteSpace(databaseName))
            {
                throw new InvalidOperationException("请先选择数据库。");
            }
        }

        private static void ValidateTable(TableItem table)
        {
            if (table == null || string.IsNullOrWhiteSpace(table.SchemaName) || string.IsNullOrWhiteSpace(table.TableName))
            {
                throw new InvalidOperationException("请先选择数据表。");
            }
        }

        private static string BuildMasterConnectionString(SqlServerConnectionOptions options)
        {
            var builder = new MySqlConnectionStringBuilder
            {
                Server = options.ServerAddress?.Trim(),
                Port = (uint)NormalizePort(options.Port, 3306),
                Database = "mysql",
                UserID = options.Username?.Trim(),
                Password = options.Password ?? string.Empty,
                ConnectionTimeout = 10,
                ApplicationName = "MyTools"
            };
            return builder.ConnectionString;
        }

        private static string BuildDatabaseConnectionString(SqlServerConnectionOptions options, string databaseName)
        {
            var builder = new MySqlConnectionStringBuilder(BuildMasterConnectionString(options))
            {
                Database = databaseName
            };
            return builder.ConnectionString;
        }

        private static int NormalizePort(string rawPort, int defaultPort)
        {
            if (!string.IsNullOrWhiteSpace(rawPort) && int.TryParse(rawPort.Trim(), out var parsedPort) && parsedPort > 0)
            {
                return parsedPort;
            }

            return defaultPort;
        }
    }
}
