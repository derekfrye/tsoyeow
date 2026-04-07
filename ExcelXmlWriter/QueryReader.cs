using System.Data;
using System.Globalization;
using System.Xml;
using Microsoft.Data.SqlClient;

namespace ExcelXmlWriter
{
    internal sealed class QueryReader : IDataReader
    {
        readonly bool fromFile;
        readonly SqlCommand command;
        readonly SqlConnection connection;
        readonly DataSet dataSet;

        IDataReader? reader;
        int tableIndex;
        int totalRecordsRead;
        int currentResult = 1;

        internal QueryReader(string query, int queryTimeout, bool queryIsFile, string connStr)
        {
            fromFile = queryIsFile;

            if (!queryIsFile)
            {
                connection = new SqlConnection(connStr);
                command = new SqlCommand(query, connection)
                {
                    CommandType = CommandType.Text,
                    CommandTimeout = queryTimeout,
                };
                dataSet = new DataSet();
                return;
            }

            dataSet = new DataSet
            {
                Locale = CultureInfo.GetCultureInfo("en-US"),
            };

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            };

            using var xmlReader = XmlReader.Create(query, settings);
            dataSet.ReadXml(xmlReader);

            command = new SqlCommand();
            connection = new SqlConnection();
        }

        internal int CurrentResult => currentResult;

        internal int TotalRecordsRead => totalRecordsRead;

        internal bool HasActiveResultSet => reader is not null;

        public void OpenReader()
        {
            if (fromFile)
            {
                if (dataSet.Tables.Count == 0)
                {
                    reader = null;
                    return;
                }

                tableIndex = 0;
                reader = dataSet.Tables[tableIndex].CreateDataReader();
                currentResult = 1;
                return;
            }

            if (reader is not null && !reader.IsClosed)
            {
                return;
            }

            try
            {
                if (command.Connection.State == ConnectionState.Closed)
                {
                    command.Connection.Open();
                }

                reader = command.ExecuteReader(CommandBehavior.CloseConnection);
                currentResult = 1;
            }
            catch
            {
                if (command.Connection.State == ConnectionState.Open)
                {
                    command.Connection.Close();
                }

                throw;
            }
        }

        public void CloseReader()
        {
            if (reader is not null && !reader.IsClosed)
            {
                reader.Close();
            }

            if (!fromFile && command.Connection.State == ConnectionState.Open)
            {
                command.Connection.Close();
            }
        }

        internal void Cancel()
        {
            if (!fromFile)
            {
                command.Cancel();
            }
        }

        IDataReader EnsureReader()
        {
            return reader ?? throw new InvalidOperationException("The query reader has not been opened.");
        }

        public object this[int i] => EnsureReader()[i];

        public object this[string name] => EnsureReader()[name];

        public int Depth => EnsureReader().Depth;

        public bool IsClosed => reader?.IsClosed ?? true;

        public int RecordsAffected => EnsureReader().RecordsAffected;

        public int FieldCount => EnsureReader().FieldCount;

        public void Close()
        {
            CloseReader();
        }

        public void Dispose()
        {
            reader?.Dispose();
            command.Dispose();
            connection.Dispose();
            dataSet.Dispose();
        }

        public bool NextResult()
        {
            if (reader is null)
            {
                return false;
            }

            if (fromFile)
            {
                tableIndex++;
                if (tableIndex >= dataSet.Tables.Count)
                {
                    return false;
                }

                reader.Dispose();
                reader = dataSet.Tables[tableIndex].CreateDataReader();
                currentResult++;
                return true;
            }

            if (reader.NextResult())
            {
                currentResult++;
                return true;
            }

            return false;
        }

        public bool Read()
        {
            var hasRow = EnsureReader().Read();
            if (hasRow)
            {
                totalRecordsRead++;
            }

            return hasRow;
        }

        public DataTable? GetSchemaTable()
        {
            return EnsureReader().GetSchemaTable();
        }

        public bool GetBoolean(int i) => EnsureReader().GetBoolean(i);

        public byte GetByte(int i) => EnsureReader().GetByte(i);

        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length)
            => EnsureReader().GetBytes(i, fieldOffset, buffer, bufferoffset, length);

        public char GetChar(int i) => EnsureReader().GetChar(i);

        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length)
            => EnsureReader().GetChars(i, fieldoffset, buffer, bufferoffset, length);

        public IDataReader GetData(int i) => EnsureReader().GetData(i);

        public string GetDataTypeName(int i) => EnsureReader().GetDataTypeName(i);

        public DateTime GetDateTime(int i) => EnsureReader().GetDateTime(i);

        public decimal GetDecimal(int i) => EnsureReader().GetDecimal(i);

        public double GetDouble(int i) => EnsureReader().GetDouble(i);

        public Type GetFieldType(int i) => EnsureReader().GetFieldType(i);

        public float GetFloat(int i) => EnsureReader().GetFloat(i);

        public Guid GetGuid(int i) => EnsureReader().GetGuid(i);

        public short GetInt16(int i) => EnsureReader().GetInt16(i);

        public int GetInt32(int i) => EnsureReader().GetInt32(i);

        public long GetInt64(int i) => EnsureReader().GetInt64(i);

        public string GetName(int i) => EnsureReader().GetName(i);

        public int GetOrdinal(string name) => EnsureReader().GetOrdinal(name);

        public string GetString(int i) => EnsureReader().GetString(i);

        public object GetValue(int i) => EnsureReader().GetValue(i);

        public int GetValues(object[] values) => EnsureReader().GetValues(values);

        public bool IsDBNull(int i) => EnsureReader().IsDBNull(i);
    }
}
