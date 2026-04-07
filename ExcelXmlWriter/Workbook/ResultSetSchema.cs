using System.Data;

namespace ExcelXmlWriter.Workbook
{
    internal sealed class ResultSetSchema
    {
        ResultSetSchema(DataTable schemaTable, IReadOnlySet<string> columnNames)
        {
            SchemaTable = schemaTable;
            ColumnNames = columnNames;
        }

        internal DataTable SchemaTable { get; }

        internal DataRowCollection Rows => SchemaTable.Rows;

        internal IReadOnlySet<string> ColumnNames { get; }

        internal bool ContainsColumn(string columnName)
        {
            return ColumnNames.Contains(columnName);
        }

        internal static ResultSetSchema? Create(DataTable? schemaTable)
        {
            if (schemaTable is null)
            {
                return null;
            }

            var columnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in schemaTable.Rows)
            {
                var columnName = row["ColumnName"]?.ToString();
                if (!string.IsNullOrEmpty(columnName))
                {
                    columnNames.Add(columnName);
                }
            }

            return new ResultSetSchema(schemaTable, columnNames);
        }
    }
}
