using System.Data;

namespace ExcelXmlWriter.Tests.TestData;

internal static class TestDataSetFactory
{
    internal static string CreateXmlFile(params (string TableName, IReadOnlyList<string> Columns, IReadOnlyList<object?[]> Rows)[] tables)
    {
        var dataSet = new DataSet("NewDataSet");
        foreach (var (tableName, columns, rows) in tables)
        {
            var table = new DataTable(tableName);
            foreach (var column in columns)
            {
                table.Columns.Add(column, typeof(string));
            }

            foreach (var row in rows)
            {
                table.Rows.Add(row);
            }

            dataSet.Tables.Add(table);
        }

        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xml");
        dataSet.WriteXml(path, XmlWriteMode.WriteSchema);
        return path;
    }
}
