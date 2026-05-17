using System;

namespace MyTools.Services
{
    public class SqlServerConnectionOptions
    {
        public SqlProviderKind ProviderKind { get; set; } = SqlProviderKind.SqlServer;
        public string ServerAddress { get; set; }
        public string Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    public class DatabaseItem
    {
        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }

    public class TableItem
    {
        public string SchemaName { get; set; }
        public string TableName { get; set; }

        public string DisplayName
        {
            get { return $"{SchemaName}.{TableName}"; }
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }

    public class ExportResult
    {
        public long RowCount { get; set; }
        public string FilePath { get; set; }
        public long FileSizeBytes { get; set; }
        public TimeSpan Duration { get; set; }
    }

    public class SqlExportProgress
    {
        public string Stage { get; set; }
        public long ProcessedRows { get; set; }
        public long? TotalRows { get; set; }
        public TimeSpan Elapsed { get; set; }
        public long FileSizeBytes { get; set; }
    }
}
