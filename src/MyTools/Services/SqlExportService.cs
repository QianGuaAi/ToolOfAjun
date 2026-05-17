using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MyTools.Services
{
    public static class SqlExportService
    {
        public const int ExcelWorksheetRowLimit = 1048576;

        public static async Task TestConnectionAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            AppLogService.Information("Testing SQL Server connection to {ServerAddress}:{Port}", options.ServerAddress, NormalizePort(options.Port));

            using (var connection = new SqlConnection(BuildMasterConnectionString(options)))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
            }
        }

        public static async Task<List<DatabaseItem>> GetDatabasesAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);

            try
            {
                var databases = await GetDatabasesFromCatalogViewAsync(options, cancellationToken).ConfigureAwait(false);
                AppLogService.Information("Loaded {DatabaseCount} SQL Server databases from {ServerAddress}", databases.Count, options.ServerAddress);
                return databases;
            }
            catch (SqlException ex) when (IsInvalidSysDatabasesError(ex))
            {
                AppLogService.Information("sys.databases is unavailable on {ServerAddress}; falling back to sp_databases.", options.ServerAddress);
                var databases = await GetDatabasesFromStoredProcedureAsync(options, cancellationToken).ConfigureAwait(false);
                AppLogService.Information("Loaded {DatabaseCount} SQL Server databases from {ServerAddress}", databases.Count, options.ServerAddress);
                return databases;
            }
        }

        private static async Task<List<DatabaseItem>> GetDatabasesFromCatalogViewAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            const string sql = @"
SELECT name
FROM sys.databases
WHERE state = 0
  AND HAS_DBACCESS(name) = 1
ORDER BY name;";

            var databases = new List<DatabaseItem>();
            using (var connection = new SqlConnection(BuildMasterConnectionString(options)))
            using (var command = new SqlCommand(sql, connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        databases.Add(new DatabaseItem
                        {
                            Name = reader.GetString(0)
                        });
                    }
                }
            }

            return databases;
        }

        private static async Task<List<DatabaseItem>> GetDatabasesFromStoredProcedureAsync(SqlServerConnectionOptions options, CancellationToken cancellationToken)
        {
            var databaseNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var connection = new SqlConnection(BuildMasterConnectionString(options)))
            using (var command = new SqlCommand("sp_databases", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        var databaseName = Convert.ToString(reader["DATABASE_NAME"], CultureInfo.InvariantCulture);
                        if (!string.IsNullOrWhiteSpace(databaseName))
                        {
                            databaseNames.Add(databaseName);
                        }
                    }
                }
            }

            var databases = new List<DatabaseItem>();
            foreach (var databaseName in databaseNames)
            {
                databases.Add(new DatabaseItem
                {
                    Name = databaseName
                });
            }

            databases.Sort((left, right) => string.Compare(left.Name, right.Name, StringComparison.CurrentCultureIgnoreCase));
            return databases;
        }

        public static async Task<List<TableItem>> GetTablesAsync(SqlServerConnectionOptions options, string databaseName, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);

            const string sql = @"
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
ORDER BY TABLE_SCHEMA, TABLE_NAME;";

            var tables = new List<TableItem>();
            using (var connection = new SqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new SqlCommand(sql, connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        tables.Add(new TableItem
                        {
                            SchemaName = reader.GetString(0),
                            TableName = reader.GetString(1)
                        });
                    }
                }
            }

            AppLogService.Information("Loaded {TableCount} tables from {DatabaseName}", tables.Count, databaseName);
            return tables;
        }

        public static async Task<long> GetTableRowCountAsync(SqlServerConnectionOptions options, string databaseName, TableItem table, CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);
            ValidateTable(table);

            var sql = $"SELECT COUNT_BIG(1) FROM {GetQualifiedTableName(table)};";

            using (var connection = new SqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new SqlCommand(sql, connection))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                var result = await command.ExecuteScalarAsync(cancellationToken).ConfigureAwait(false);
                return Convert.ToInt64(result, CultureInfo.InvariantCulture);
            }
        }

        public static async Task<ExportResult> ExportTableAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            TableItem table,
            string filePath,
            CancellationToken cancellationToken,
            IProgress<SqlExportProgress> progress = null)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);
            ValidateTable(table);

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("导出文件路径不能为空。", nameof(filePath));
            }

            var stopwatch = Stopwatch.StartNew();
            ReportProgress(progress, "正在统计行数", 0, null, stopwatch, filePath);
            var rowCount = await GetTableRowCountAsync(options, databaseName, table, cancellationToken).ConfigureAwait(false);
            ReportProgress(progress, "正在读取并写入", 0, rowCount, stopwatch, filePath);

            Directory.CreateDirectory(Path.GetDirectoryName(filePath) ?? AppDomain.CurrentDomain.BaseDirectory);

            var selectSql = $"SELECT * FROM {GetQualifiedTableName(table)};";
            AppLogService.Information(
                "Starting SQL export from {ServerAddress}, database {DatabaseName}, table {SchemaName}.{TableName}",
                options.ServerAddress,
                databaseName,
                table.SchemaName,
                table.TableName);

            using (var connection = new SqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new SqlCommand(selectSql, connection))
            {
                command.CommandTimeout = 0;
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

                using (var reader = await command.ExecuteReaderAsync(CommandBehavior.SequentialAccess | CommandBehavior.CloseConnection, cancellationToken).ConfigureAwait(false))
                {
                    await XlsxStreamWriter.WriteAsync(
                            filePath,
                            table.DisplayName,
                            reader,
                            cancellationToken,
                            rowCount,
                            progress,
                            stopwatch)
                        .ConfigureAwait(false);
                }
            }

            stopwatch.Stop();
            var fileSizeBytes = GetFileSize(filePath);
            ReportProgress(progress, "导出完成", rowCount, rowCount, stopwatch, filePath);
            AppLogService.Information(
                "SQL export completed for {DatabaseName}.{SchemaName}.{TableName} with {RowCount} rows in {ElapsedMs} ms, file size {FileSizeBytes} bytes",
                databaseName,
                table.SchemaName,
                table.TableName,
                rowCount,
                stopwatch.ElapsedMilliseconds,
                fileSizeBytes);

            return new ExportResult
            {
                FilePath = filePath,
                RowCount = rowCount,
                FileSizeBytes = fileSizeBytes,
                Duration = stopwatch.Elapsed
            };
        }

        public static string BuildDefaultFileName(string serverAddress, string databaseName, TableItem table)
        {
            var serverPart = SanitizeFileNamePart(serverAddress);
            var databasePart = SanitizeFileNamePart(databaseName);
            var tablePart = SanitizeFileNamePart(table.DisplayName);
            return $"{serverPart}_{databasePart}_{tablePart}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        }

        public static async Task<DataTable> ExecuteQueryAsync(
            SqlServerConnectionOptions options,
            string databaseName,
            string sql,
            CancellationToken cancellationToken)
        {
            ValidateConnectionOptions(options);
            ValidateDatabaseName(databaseName);

            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException("请输入 SQL 语句。");

            AppLogService.Information(
                "Executing SQL query on {ServerAddress}/{DatabaseName}",
                options.ServerAddress,
                databaseName);

            var table = new DataTable();
            using (var connection = new SqlConnection(BuildDatabaseConnectionString(options, databaseName)))
            using (var command = new SqlCommand(sql, connection) { CommandTimeout = 120 })
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                {
                    table.Load(reader);
                }
            }

            AppLogService.Information(
                "SQL query returned {RowCount} rows and {ColCount} columns",
                table.Rows.Count,
                table.Columns.Count);

            return table;
        }

        public static async Task<ExportResult> ExportDataTableAsync(
            DataTable dataTable,
            string sheetName,
            string filePath,
            CancellationToken cancellationToken,
            IProgress<SqlExportProgress> progress = null)
        {
            if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("导出文件路径不能为空。", nameof(filePath));

            var stopwatch = Stopwatch.StartNew();
            Directory.CreateDirectory(
                Path.GetDirectoryName(filePath) ?? AppDomain.CurrentDomain.BaseDirectory);
            ReportProgress(progress, "正在写入 Excel", 0, dataTable.Rows.Count, stopwatch, filePath);

            await XlsxStreamWriter
                .WriteFromDataTableAsync(filePath, sheetName, dataTable, cancellationToken, progress, stopwatch)
                .ConfigureAwait(false);

            stopwatch.Stop();
            var fileSizeBytes = GetFileSize(filePath);
            ReportProgress(progress, "导出完成", dataTable.Rows.Count, dataTable.Rows.Count, stopwatch, filePath);
            AppLogService.Information(
                "DataTable exported to {FilePath} with {RowCount} rows in {ElapsedMs} ms, file size {FileSizeBytes} bytes",
                filePath,
                dataTable.Rows.Count,
                stopwatch.ElapsedMilliseconds,
                fileSizeBytes);

            return new ExportResult
            {
                FilePath = filePath,
                RowCount = dataTable.Rows.Count,
                FileSizeBytes = fileSizeBytes,
                Duration = stopwatch.Elapsed
            };
        }

        public static async Task<ExportResult> ExportDataTableToCsvAsync(
            DataTable dataTable,
            string filePath,
            CancellationToken cancellationToken,
            IProgress<SqlExportProgress> progress = null)
        {
            if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("导出文件路径不能为空。", nameof(filePath));

            var stopwatch = Stopwatch.StartNew();
            Directory.CreateDirectory(
                Path.GetDirectoryName(filePath) ?? AppDomain.CurrentDomain.BaseDirectory);
            ReportProgress(progress, "正在写入 CSV", 0, dataTable.Rows.Count, stopwatch, filePath);

            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                {
                    if (columnIndex > 0) await writer.WriteAsync(",").ConfigureAwait(false);
                    await writer.WriteAsync(EscapeCsvValue(dataTable.Columns[columnIndex].ColumnName)).ConfigureAwait(false);
                }

                await writer.WriteLineAsync().ConfigureAwait(false);

                long processedRows = 0;
                foreach (DataRow row in dataTable.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    processedRows++;
                    for (var columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                    {
                        if (columnIndex > 0) await writer.WriteAsync(",").ConfigureAwait(false);
                        await writer.WriteAsync(EscapeCsvValue(FormatCsvValue(row[columnIndex]))).ConfigureAwait(false);
                    }

                    await writer.WriteLineAsync().ConfigureAwait(false);
                    if (processedRows == 1 || processedRows % 1000 == 0)
                    {
                        ReportProgress(progress, "正在写入 CSV", processedRows, dataTable.Rows.Count, stopwatch, filePath);
                    }
                }
            }

            stopwatch.Stop();
            var fileSizeBytes = GetFileSize(filePath);
            ReportProgress(progress, "CSV 导出完成", dataTable.Rows.Count, dataTable.Rows.Count, stopwatch, filePath);
            AppLogService.Information(
                "DataTable exported to CSV {FilePath} with {RowCount} rows in {ElapsedMs} ms, file size {FileSizeBytes} bytes",
                filePath,
                dataTable.Rows.Count,
                stopwatch.ElapsedMilliseconds,
                fileSizeBytes);

            return new ExportResult
            {
                FilePath = filePath,
                RowCount = dataTable.Rows.Count,
                FileSizeBytes = fileSizeBytes,
                Duration = stopwatch.Elapsed
            };
        }

        public static void ReportProgress(
            IProgress<SqlExportProgress> progress,
            string stage,
            long processedRows,
            long? totalRows,
            Stopwatch stopwatch,
            string filePath)
        {
            if (progress == null)
            {
                return;
            }

            progress.Report(new SqlExportProgress
            {
                Stage = stage,
                ProcessedRows = processedRows,
                TotalRows = totalRows,
                Elapsed = stopwatch?.Elapsed ?? TimeSpan.Zero,
                FileSizeBytes = GetFileSize(filePath)
            });
        }

        private static long GetFileSize(string filePath)
        {
            try
            {
                return string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)
                    ? 0
                    : new FileInfo(filePath).Length;
            }
            catch
            {
                return 0;
            }
        }

        private static string FormatCsvValue(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            if (value is IFormattable formattable)
            {
                return formattable.ToString(null, CultureInfo.CurrentCulture);
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
        }

        private static string EscapeCsvValue(string value)
        {
            value = value ?? string.Empty;
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0)
            {
                return value;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static void ValidateConnectionOptions(SqlServerConnectionOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.ServerAddress))
            {
                throw new InvalidOperationException("请输入 SQL Server 地址。");
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
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = BuildDataSource(options.ServerAddress, options.Port),
                InitialCatalog = "master",
                UserID = options.Username,
                Password = options.Password ?? string.Empty,
                PersistSecurityInfo = false,
                IntegratedSecurity = false,
                ConnectTimeout = 10,
                ApplicationName = "MyTools"
            };

            return builder.ConnectionString;
        }

        private static string BuildDatabaseConnectionString(SqlServerConnectionOptions options, string databaseName)
        {
            var builder = new SqlConnectionStringBuilder(BuildMasterConnectionString(options))
            {
                InitialCatalog = databaseName
            };

            return builder.ConnectionString;
        }

        private static bool IsInvalidSysDatabasesError(SqlException exception)
        {
            foreach (SqlError error in exception.Errors)
            {
                if (error.Number == 208 && error.Message.IndexOf("sys.databases", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string BuildDataSource(string serverAddress, string port)
        {
            var normalizedPort = NormalizePort(port);
            return string.IsNullOrWhiteSpace(normalizedPort) ? serverAddress.Trim() : $"{serverAddress.Trim()},{normalizedPort}";
        }

        private static string NormalizePort(string port)
        {
            return string.IsNullOrWhiteSpace(port) ? string.Empty : port.Trim();
        }

        private static string GetQualifiedTableName(TableItem table)
        {
            return $"{EscapeSqlIdentifier(table.SchemaName)}.{EscapeSqlIdentifier(table.TableName)}";
        }

        private static string EscapeSqlIdentifier(string value)
        {
            return "[" + value.Replace("]", "]]") + "]";
        }

        private static string SanitizeFileNamePart(string value)
        {
            var sanitized = value ?? "export";
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                sanitized = sanitized.Replace(invalidChar, '_');
            }

            return sanitized.Replace('.', '_').Trim('_');
        }

        private static class XlsxStreamWriter
        {
            private const int DateStyleIndex = 1;
            private const int DateTimeStyleIndex = 2;
            private const int MaxDataRowsPerWorksheet = ExcelWorksheetRowLimit - 1;

            public static async Task WriteAsync(
                string filePath,
                string sheetName,
                SqlDataReader reader,
                CancellationToken cancellationToken,
                long? totalRows = null,
                IProgress<SqlExportProgress> progress = null,
                Stopwatch stopwatch = null)
            {
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, leaveOpen: false))
                {
                    var sheetCount = GetWorksheetCount(totalRows);
                    WriteEntry(archive, "[Content_Types].xml", BuildContentTypesXml(sheetCount));
                    WriteEntry(archive, "_rels/.rels", BuildRootRelsXml());
                    WriteEntry(archive, "docProps/app.xml", BuildAppXml());
                    WriteEntry(archive, "docProps/core.xml", BuildCoreXml());
                    WriteEntry(archive, "xl/workbook.xml", BuildWorkbookXml(sheetName, sheetCount));
                    WriteEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml(sheetCount));
                    WriteEntry(archive, "xl/styles.xml", BuildStylesXml());

                    await WriteWorksheetsAsync(archive, reader, cancellationToken, totalRows, progress, stopwatch, filePath, sheetCount)
                        .ConfigureAwait(false);
                }
            }

            private static void WriteEntry(ZipArchive archive, string path, string content)
            {
                var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
                using (var stream = entry.Open())
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(content);
                }
            }

            private static int GetWorksheetCount(long? totalRows)
            {
                if (!totalRows.HasValue || totalRows.Value <= 0)
                {
                    return 1;
                }

                return (int)Math.Max(1, (totalRows.Value + MaxDataRowsPerWorksheet - 1) / MaxDataRowsPerWorksheet);
            }

            private static async Task WriteWorksheetsAsync(
                ZipArchive archive,
                SqlDataReader reader,
                CancellationToken cancellationToken,
                long? totalRows,
                IProgress<SqlExportProgress> progress,
                Stopwatch stopwatch,
                string filePath,
                int sheetCount)
            {
                long processedRows = 0;
                for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
                {
                    var worksheetEntry = archive.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml", CompressionLevel.Fastest);
                    using (var stream = worksheetEntry.Open())
                    using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
                    {
                        Async = true,
                        Encoding = new UTF8Encoding(false),
                        CloseOutput = false
                    }))
                    {
                        var writtenRows = await WriteWorksheetPartAsync(
                                writer,
                                reader,
                                cancellationToken,
                                totalRows,
                                progress,
                                stopwatch,
                                filePath,
                                sheetIndex,
                                processedRows)
                            .ConfigureAwait(false);
                        processedRows += writtenRows;
                        if (writtenRows < MaxDataRowsPerWorksheet)
                        {
                            break;
                        }
                    }
                }
            }

            private static async Task WriteWorksheetAsync(
                XmlWriter writer,
                SqlDataReader reader,
                CancellationToken cancellationToken,
                long? totalRows,
                IProgress<SqlExportProgress> progress,
                Stopwatch stopwatch,
                string filePath)
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                writer.WriteStartElement("sheetData");

                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", "1");
                for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                {
                    WriteInlineStringCell(writer, GetCellReference(columnIndex, 1), reader.GetName(columnIndex));
                }

                writer.WriteEndElement();

                var rowIndex = 1;
                long processedRows = 0;
                while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    rowIndex++;
                    processedRows++;

                    writer.WriteStartElement("row");
                    writer.WriteAttributeString("r", rowIndex.ToString(CultureInfo.InvariantCulture));
                    for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                    {
                        WriteCell(writer, reader, columnIndex, rowIndex);
                    }

                    writer.WriteEndElement();
                    if (processedRows == 1 || processedRows % 1000 == 0)
                    {
                        SqlExportService.ReportProgress(progress, "正在读取并写入", processedRows, totalRows, stopwatch, filePath);
                    }
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                await writer.FlushAsync().ConfigureAwait(false);
            }

            private static async Task<long> WriteWorksheetPartAsync(
                XmlWriter writer,
                SqlDataReader reader,
                CancellationToken cancellationToken,
                long? totalRows,
                IProgress<SqlExportProgress> progress,
                Stopwatch stopwatch,
                string filePath,
                int sheetIndex,
                long baseProcessedRows)
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                writer.WriteStartElement("sheetData");

                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", "1");
                for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                {
                    WriteInlineStringCell(writer, GetCellReference(columnIndex, 1), reader.GetName(columnIndex));
                }

                writer.WriteEndElement();

                var rowIndex = 1;
                long processedRows = 0;
                while (processedRows < MaxDataRowsPerWorksheet && await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    rowIndex++;
                    processedRows++;

                    writer.WriteStartElement("row");
                    writer.WriteAttributeString("r", rowIndex.ToString(CultureInfo.InvariantCulture));
                    for (var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
                    {
                        WriteCell(writer, reader, columnIndex, rowIndex);
                    }

                    writer.WriteEndElement();
                    if (processedRows == 1 || processedRows % 1000 == 0)
                    {
                        var stage = sheetIndex > 1 ? $"正在写入 Sheet {sheetIndex}" : "正在读取并写入";
                        SqlExportService.ReportProgress(progress, stage, baseProcessedRows + processedRows, totalRows, stopwatch, filePath);
                    }
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                await writer.FlushAsync().ConfigureAwait(false);
                return processedRows;
            }

            private static void WriteCell(XmlWriter writer, IDataRecord record, int columnIndex, int rowIndex)
            {
                WriteValueCell(
                    writer,
                    GetCellReference(columnIndex, rowIndex),
                    record.IsDBNull(columnIndex) ? null : record.GetValue(columnIndex));
            }

            private static void WriteValueCell(XmlWriter writer, string cellReference, object value)
            {
                if (value == null)
                {
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", cellReference);
                    writer.WriteEndElement();
                    return;
                }

                switch (value)
                {
                    case bool boolValue:
                        writer.WriteStartElement("c");
                        writer.WriteAttributeString("r", cellReference);
                        writer.WriteAttributeString("t", "b");
                        writer.WriteElementString("v", boolValue ? "1" : "0");
                        writer.WriteEndElement();
                        return;
                    case byte byteValue:
                        WriteNumberCell(writer, cellReference, byteValue.ToString(CultureInfo.InvariantCulture), null);
                        return;
                    case short shortValue:
                        WriteNumberCell(writer, cellReference, shortValue.ToString(CultureInfo.InvariantCulture), null);
                        return;
                    case int intValue:
                        WriteNumberCell(writer, cellReference, intValue.ToString(CultureInfo.InvariantCulture), null);
                        return;
                    case long longValue:
                        WriteNumberCell(writer, cellReference, longValue.ToString(CultureInfo.InvariantCulture), null);
                        return;
                    case decimal decimalValue:
                        WriteNumberCell(writer, cellReference, decimalValue.ToString(CultureInfo.InvariantCulture), null);
                        return;
                    case float floatValue:
                        WriteNumberCell(writer, cellReference, floatValue.ToString("R", CultureInfo.InvariantCulture), null);
                        return;
                    case double doubleValue:
                        WriteNumberCell(writer, cellReference, doubleValue.ToString("R", CultureInfo.InvariantCulture), null);
                        return;
                    case DateTime dateTimeValue:
                        var styleIndex = dateTimeValue.TimeOfDay == TimeSpan.Zero ? DateStyleIndex : DateTimeStyleIndex;
                        WriteNumberCell(writer, cellReference, dateTimeValue.ToOADate().ToString(CultureInfo.InvariantCulture), styleIndex);
                        return;
                    case Guid guidValue:
                        WriteInlineStringCell(writer, cellReference, guidValue.ToString());
                        return;
                    case byte[] binaryValue:
                        WriteInlineStringCell(writer, cellReference, Convert.ToBase64String(binaryValue));
                        return;
                    default:
                        WriteInlineStringCell(writer, cellReference, Convert.ToString(value, CultureInfo.CurrentCulture));
                        return;
                }
            }

            public static async Task WriteFromDataTableAsync(
                string filePath,
                string sheetName,
                DataTable dataTable,
                CancellationToken cancellationToken,
                IProgress<SqlExportProgress> progress = null,
                Stopwatch stopwatch = null)
            {
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, leaveOpen: false))
                {
                    WriteEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
                    WriteEntry(archive, "_rels/.rels", BuildRootRelsXml());
                    WriteEntry(archive, "docProps/app.xml", BuildAppXml());
                    WriteEntry(archive, "docProps/core.xml", BuildCoreXml());
                    WriteEntry(archive, "xl/workbook.xml", BuildWorkbookXml(sheetName));
                    WriteEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelsXml());
                    WriteEntry(archive, "xl/styles.xml", BuildStylesXml());

                    var worksheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
                    using (var stream = worksheetEntry.Open())
                    using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
                    {
                        Async = true,
                        Encoding = new UTF8Encoding(false),
                        CloseOutput = false
                    }))
                    {
                        await WriteWorksheetFromDataTableAsync(writer, dataTable, cancellationToken, progress, stopwatch, filePath)
                            .ConfigureAwait(false);
                    }
                }
            }

            private static async Task WriteWorksheetFromDataTableAsync(
                XmlWriter writer,
                DataTable dataTable,
                CancellationToken cancellationToken,
                IProgress<SqlExportProgress> progress,
                Stopwatch stopwatch,
                string filePath)
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                writer.WriteStartElement("sheetData");

                writer.WriteStartElement("row");
                writer.WriteAttributeString("r", "1");
                for (var colIdx = 0; colIdx < dataTable.Columns.Count; colIdx++)
                {
                    WriteInlineStringCell(writer, GetCellReference(colIdx, 1), dataTable.Columns[colIdx].ColumnName);
                }
                writer.WriteEndElement();

                var rowIndex = 1;
                long processedRows = 0;
                foreach (DataRow row in dataTable.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    rowIndex++;
                    processedRows++;
                    writer.WriteStartElement("row");
                    writer.WriteAttributeString("r", rowIndex.ToString(CultureInfo.InvariantCulture));
                    for (var colIdx = 0; colIdx < dataTable.Columns.Count; colIdx++)
                    {
                        WriteValueCell(
                            writer,
                            GetCellReference(colIdx, rowIndex),
                            row.IsNull(colIdx) ? null : row[colIdx]);
                    }
                    writer.WriteEndElement();
                    if (processedRows == 1 || processedRows % 1000 == 0)
                    {
                        SqlExportService.ReportProgress(progress, "正在写入 Excel", processedRows, dataTable.Rows.Count, stopwatch, filePath);
                    }
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                await writer.FlushAsync().ConfigureAwait(false);
            }

            private static void WriteInlineStringCell(XmlWriter writer, string cellReference, string text)
            {
                writer.WriteStartElement("c");
                writer.WriteAttributeString("r", cellReference);
                writer.WriteAttributeString("t", "inlineStr");
                writer.WriteStartElement("is");
                writer.WriteStartElement("t");
                if (!string.IsNullOrEmpty(text) && (text[0] == ' ' || text[text.Length - 1] == ' '))
                {
                    writer.WriteAttributeString("xml", "space", null, "preserve");
                }

                writer.WriteString(SanitizeText(text));
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            private static void WriteNumberCell(XmlWriter writer, string cellReference, string value, int? styleIndex)
            {
                writer.WriteStartElement("c");
                writer.WriteAttributeString("r", cellReference);
                if (styleIndex.HasValue)
                {
                    writer.WriteAttributeString("s", styleIndex.Value.ToString(CultureInfo.InvariantCulture));
                }

                writer.WriteElementString("v", value);
                writer.WriteEndElement();
            }

            private static string GetCellReference(int columnIndex, int rowIndex)
            {
                return GetColumnName(columnIndex + 1) + rowIndex.ToString(CultureInfo.InvariantCulture);
            }

            private static string GetColumnName(int columnNumber)
            {
                var builder = new StringBuilder();
                while (columnNumber > 0)
                {
                    var modulo = (columnNumber - 1) % 26;
                    builder.Insert(0, Convert.ToChar('A' + modulo));
                    columnNumber = (columnNumber - modulo) / 26;
                }

                return builder.ToString();
            }

            private static string SanitizeText(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return string.Empty;
                }

                var builder = new StringBuilder(value.Length);
                foreach (var ch in value)
                {
                    if (XmlConvert.IsXmlChar(ch))
                    {
                        builder.Append(ch);
                    }
                }

                return builder.ToString();
            }

            private static string BuildContentTypesXml()
            {
                return BuildContentTypesXml(1);
            }

            private static string BuildContentTypesXml(int sheetCount)
            {
                var worksheetOverrides = new StringBuilder();
                for (var i = 1; i <= Math.Max(1, sheetCount); i++)
                {
                    worksheetOverrides.Append("  <Override PartName=\"/xl/worksheets/sheet")
                        .Append(i.ToString(CultureInfo.InvariantCulture))
                        .AppendLine(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />");
                }

                return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
  <Default Extension=""xml"" ContentType=""application/xml"" />
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" />
{worksheetOverrides.ToString().TrimEnd()}
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
  <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.core-properties+xml"" />
  <Override PartName=""/docProps/app.xml"" ContentType=""application/vnd.openxmlformats-officedocument.extended-properties+xml"" />
</Types>";
            }

            private static string BuildRootRelsXml()
            {
                return @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" />
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml"" />
  <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml"" />
</Relationships>";
            }

            private static string BuildAppXml()
            {
                return @"<?xml version=""1.0"" encoding=""utf-8""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties""
            xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
  <Application>MyTools</Application>
</Properties>";
            }

            private static string BuildCoreXml()
            {
                var created = XmlConvert.ToString(DateTime.UtcNow, XmlDateTimeSerializationMode.Utc);
                return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""
                   xmlns:dc=""http://purl.org/dc/elements/1.1/""
                   xmlns:dcterms=""http://purl.org/dc/terms/""
                   xmlns:dcmitype=""http://purl.org/dc/dcmitype/""
                   xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <dc:creator>MyTools</dc:creator>
  <cp:lastModifiedBy>MyTools</cp:lastModifiedBy>
  <dcterms:created xsi:type=""dcterms:W3CDTF"">{created}</dcterms:created>
  <dcterms:modified xsi:type=""dcterms:W3CDTF"">{created}</dcterms:modified>
</cp:coreProperties>";
            }

            private static string BuildWorkbookXml(string sheetName)
            {
                return BuildWorkbookXml(sheetName, 1);
            }

            private static string BuildWorkbookXml(string sheetName, int sheetCount)
            {
                var sheets = new StringBuilder();
                var baseName = NormalizeSheetName(sheetName);
                for (var i = 1; i <= Math.Max(1, sheetCount); i++)
                {
                    var name = Math.Max(1, sheetCount) == 1 ? baseName : NormalizeSheetName(baseName + " " + i.ToString(CultureInfo.InvariantCulture));
                    sheets.Append("    <sheet name=\"")
                        .Append(SecurityElement.Escape(name))
                        .Append("\" sheetId=\"")
                        .Append(i.ToString(CultureInfo.InvariantCulture))
                        .Append("\" r:id=\"rId")
                        .Append(i.ToString(CultureInfo.InvariantCulture))
                        .AppendLine("\" />");
                }

                return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
{sheets.ToString().TrimEnd()}
  </sheets>
</workbook>";
            }

            private static string BuildWorkbookRelsXml()
            {
                return BuildWorkbookRelsXml(1);
            }

            private static string BuildWorkbookRelsXml(int sheetCount)
            {
                var relationships = new StringBuilder();
                for (var i = 1; i <= Math.Max(1, sheetCount); i++)
                {
                    relationships.Append("  <Relationship Id=\"rId")
                        .Append(i.ToString(CultureInfo.InvariantCulture))
                        .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                        .Append(i.ToString(CultureInfo.InvariantCulture))
                        .AppendLine(".xml\" />");
                }

                var styleRelId = Math.Max(1, sheetCount) + 1;
                return $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
{relationships.ToString().TrimEnd()}
  <Relationship Id=""rId{styleRelId.ToString(CultureInfo.InvariantCulture)}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml"" />
</Relationships>";
            }

            private static string BuildStylesXml()
            {
                return @"<?xml version=""1.0"" encoding=""utf-8""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <numFmts count=""2"">
    <numFmt numFmtId=""164"" formatCode=""yyyy-mm-dd"" />
    <numFmt numFmtId=""165"" formatCode=""yyyy-mm-dd hh:mm:ss"" />
  </numFmts>
  <fonts count=""1"">
    <font>
      <sz val=""11"" />
      <name val=""Calibri"" />
    </font>
  </fonts>
  <fills count=""2"">
    <fill><patternFill patternType=""none"" /></fill>
    <fill><patternFill patternType=""gray125"" /></fill>
  </fills>
  <borders count=""1"">
    <border />
  </borders>
  <cellStyleXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" />
  </cellStyleXfs>
  <cellXfs count=""3"">
    <xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" />
    <xf numFmtId=""164"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1"" />
    <xf numFmtId=""165"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1"" />
  </cellXfs>
  <cellStyles count=""1"">
    <cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
  </cellStyles>
</styleSheet>";
            }

            private static string NormalizeSheetName(string value)
            {
                var name = string.IsNullOrWhiteSpace(value) ? "Sheet1" : value.Trim();
                var invalidChars = new[] { '\\', '/', '?', '*', '[', ']', ':' };
                foreach (var invalidChar in invalidChars)
                {
                    name = name.Replace(invalidChar, '_');
                }

                return name.Length <= 31 ? name : name.Substring(0, 31);
            }
        }
    }
}
