using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace ExcelXmlQueryResults
{
    class Exceptions
    {
        [Serializable()]
        public class ConfigFileBroken : Exception
        {
            public ConfigFileBroken() : base() { throw new NotImplementedException(); }
            public ConfigFileBroken(string message) : base(message) { }
            public ConfigFileBroken(string message, Exception innerException) : base(message, innerException) { throw new NotImplementedException(); }
            protected ConfigFileBroken(SerializationInfo info, StreamingContext context) : base(info, context) { throw new NotImplementedException(); }
        }
    }
}
