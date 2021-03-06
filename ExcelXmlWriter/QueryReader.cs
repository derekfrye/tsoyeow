using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Collections;
using ExcelXmlWriter;
using System.Globalization;

namespace ExcelXmlWriter
{
    class QueryReader : IEnumerator, IDisposable, IDataReader
    {

        #region LocalFields

        IDataReader dr;
        bool fromFile;

        SqlCommand sc = new SqlCommand();
        SqlConnection scc = new SqlConnection();

        DataSet ds = new DataSet();
        int tableCount;

        readonly object dataReaderLocker = new object();
        /// <summary>
        ///lock to read/write 
        /// </summary>
        int totalRecordsRead = 0;
        readonly object currentResultLocker = new object();
        /// <summary>
        ///  lock to read/write
        /// </summary>
        int currentResult = 1;

        bool currentResultSetStillHasRecords;
        bool rowConsumed;

        #endregion

        /// <summary>
        /// 1-based result number maintained separately from the DataReader NextResult() call.
        /// </summary>
        internal int CurrentResult
        {
            get
            {
                lock (currentResultLocker)
                    return currentResult;
            }            
        }

        internal QueryReader(string query, int queryTimeout, bool queryIsFile, string connStr)
        {
            if (!queryIsFile)
            {
                this.scc = new SqlConnection(connStr);
                this.sc = new SqlCommand(query, this.scc);
                this.sc.CommandType = System.Data.CommandType.Text;
                this.sc.CommandTimeout = queryTimeout;
                currentResultSetStillHasRecords = true;
                rowConsumed = true;
            }
            else
            {
                ds = new DataSet
                {
                    Locale = CultureInfo.GetCultureInfo("en-US")
                };
                ds.ReadXml(query);
                fromFile = true;
                currentResultSetStillHasRecords = true;
            }
        }

        internal int TotalRecordsRead
        {
            get
            {
                lock (dataReaderLocker)
                    return totalRecordsRead;
            }
        }

        public void OpenReader()
        {
            if (fromFile)
                dr = ds.Tables[tableCount].CreateDataReader();
            else
            {
                if (dr == null || dr.IsClosed)
                {
                    try
                    {
                        if (sc.Connection.State == ConnectionState.Closed)
                            sc.Connection.Open();

                        //SqlDataAdapter sa = new SqlDataAdapter(sc);
                        //ds = new DataSet();
                        //sa.Fill(ds);
                        //ds.WriteXml(@"P:\ExcelXmlQueryResults\ExcelXmlQueryResults\bin\Debug\test.xml");

                        dr = sc.ExecuteReader(CommandBehavior.CloseConnection);
                    }
                    catch
                    {
                        if (sc.Connection.State == ConnectionState.Open)
                            sc.Connection.Close();

                        throw;
                    }
                }
            }
        }

        public void CloseReader()
        {
            if (!fromFile && sc.Connection.State == ConnectionState.Open)
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                sc.Connection.Close();
            }
        }

        public bool MoveToNextResultSet()
        {
            if (currentResultSetStillHasRecords)
                return currentResultSetStillHasRecords;

            tableCount++;
            if (fromFile)
            {
                if (ds.Tables.Count > tableCount)
                {
                    dr = ds.Tables[tableCount].CreateDataReader();
                    lock (currentResultLocker)
                        currentResult++;
                    currentResultSetStillHasRecords = true;
                    return true;
                }
                else
                    return false;
            }
            else
                if (dr.NextResult())
                {
                    lock (currentResultLocker)
                        currentResult++;
                    currentResultSetStillHasRecords = true;
                    return true;
                }
                else
                    return false;
        }

        public object this[int i]
        {
            get
            {
                rowConsumed = true;
                return dr[i];
            }
        }

        public object this[string i]
        {
            get
            {
                rowConsumed = true;
                return dr[i];
            }
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            rowConsumed = true;
            return dr.GetBytes(i, fieldOffset, buffer, bufferoffset, length);
        }

        public int FieldCount
        {
            get { return dr.FieldCount; }
        }

        public DataTable GetSchemaTable()
        {
            var t=dr.GetSchemaTable();
            return t;
        }

        #region IEnumerator Members

        public object Current
        {
            get { throw new NotImplementedException(); }
        }

        /// <summary>
        /// Read the next record.
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            if(!rowConsumed)
                return currentResultSetStillHasRecords;

            if (dr.Read())
            {
                lock (dataReaderLocker)
                    totalRecordsRead++;
                currentResultSetStillHasRecords = true;
                rowConsumed = false;
            }
            else
            {
                currentResultSetStillHasRecords = false;
                rowConsumed = true;
            }

            return currentResultSetStillHasRecords;
        }

        /// <summary>
        /// Call to MoveNext() will consume current row.
        /// </summary>
        public void Reset()
        {
            rowConsumed = false;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                ds.Dispose();
                sc.Dispose();
                scc.Dispose();
            }
        }


        #endregion

        #region IDataReader Members

        public void Close()
        {
            throw new NotImplementedException();
        }

        public int Depth
        {
            get { throw new NotImplementedException(); }
        }

        public bool IsClosed
        {
            get { throw new NotImplementedException(); }
        }

        public bool NextResult()
        {
            throw new NotImplementedException();
        }

        public bool Read()
        {
            throw new NotImplementedException();
        }

        public int RecordsAffected
        {
            get { throw new NotImplementedException(); }
        }

        #endregion

        #region IDataRecord Members


        public bool GetBoolean(int i)
        {
            throw new NotImplementedException();
        }

        public byte GetByte(int i)
        {
            throw new NotImplementedException();
        }

        public char GetChar(int i)
        {
            throw new NotImplementedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotImplementedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotImplementedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotImplementedException();
        }

        public DateTime GetDateTime(int i)
        {
            throw new NotImplementedException();
        }

        public decimal GetDecimal(int i)
        {
            throw new NotImplementedException();
        }

        public double GetDouble(int i)
        {
            throw new NotImplementedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotImplementedException();
        }

        public float GetFloat(int i)
        {
            throw new NotImplementedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotImplementedException();
        }

        public short GetInt16(int i)
        {
            throw new NotImplementedException();
        }

        public int GetInt32(int i)
        {
            throw new NotImplementedException();
        }

        public long GetInt64(int i)
        {
            throw new NotImplementedException();
        }

        public string GetName(int i)
        {
            throw new NotImplementedException();
        }

        public int GetOrdinal(string name)
        {
            throw new NotImplementedException();
        }

        public string GetString(int i)
        {
            throw new NotImplementedException();
        }

        public object GetValue(int i)
        {
            rowConsumed = true;
            return dr.GetValue(i);
        }

        public int GetValues(object[] values)
        {
            rowConsumed = true;
            return dr.GetValues(values);
        }

        public bool IsDBNull(int i)
        {
            rowConsumed = true;
            return dr.IsDBNull(i);
        }

        #endregion
    }
}