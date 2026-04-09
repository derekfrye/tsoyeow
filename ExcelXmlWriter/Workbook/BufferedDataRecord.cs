using System.Data;
using System.Globalization;

namespace ExcelXmlWriter.Workbook
{
    internal sealed class BufferedDataRecord : IDataRecord
    {
        readonly object?[] values;

        internal BufferedDataRecord(object?[] values)
        {
            this.values = values;
        }

        public int FieldCount => values.Length;

        public object this[int i] => GetValue(i);

        public object this[string name] => throw new NotSupportedException();

        public bool GetBoolean(int i) => Convert.ToBoolean(values[i], CultureInfo.InvariantCulture);

        public byte GetByte(int i) => Convert.ToByte(values[i], CultureInfo.InvariantCulture);

        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => throw new NotSupportedException();

        public char GetChar(int i) => Convert.ToChar(values[i], CultureInfo.InvariantCulture);

        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => throw new NotSupportedException();

        public IDataReader GetData(int i) => throw new NotSupportedException();

        public string GetDataTypeName(int i) => GetFieldType(i).Name;

        public DateTime GetDateTime(int i) => Convert.ToDateTime(values[i], CultureInfo.InvariantCulture);

        public decimal GetDecimal(int i) => Convert.ToDecimal(values[i], CultureInfo.InvariantCulture);

        public double GetDouble(int i) => Convert.ToDouble(values[i], CultureInfo.InvariantCulture);

        public Type GetFieldType(int i) => values[i]?.GetType() ?? typeof(object);

        public float GetFloat(int i) => Convert.ToSingle(values[i], CultureInfo.InvariantCulture);

        public Guid GetGuid(int i) => values[i] is Guid guid ? guid : Guid.Parse(values[i]?.ToString() ?? string.Empty);

        public short GetInt16(int i) => Convert.ToInt16(values[i], CultureInfo.InvariantCulture);

        public int GetInt32(int i) => Convert.ToInt32(values[i], CultureInfo.InvariantCulture);

        public long GetInt64(int i) => Convert.ToInt64(values[i], CultureInfo.InvariantCulture);

        public string GetName(int i) => i.ToString(CultureInfo.InvariantCulture);

        public int GetOrdinal(string name) => throw new NotSupportedException();

        public string GetString(int i) => values[i]?.ToString() ?? string.Empty;

        public object GetValue(int i) => values[i] ?? DBNull.Value;

        public int GetValues(object[] targetValues)
        {
            var length = Math.Min(targetValues.Length, values.Length);
            for (int index = 0; index < length; index++)
            {
                targetValues[index] = GetValue(index);
            }

            return length;
        }

        public bool IsDBNull(int i) => values[i] is null or DBNull;
    }
}
