using System.Text;

namespace ExcelXmlWriter.Xlsx
{
    internal sealed class StringWriterWithEncoding : StringWriter
    {
        readonly Encoding encoding;

        public StringWriterWithEncoding(Encoding encoding)
        {
            this.encoding = encoding;
        }

        public override Encoding Encoding
        {
            get
            {
                return encoding;
            }
        }
    }
}
