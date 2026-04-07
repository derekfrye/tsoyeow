namespace ExcelXmlWriter.Xlsx
{
    internal sealed class TempFileBuffer : IDisposable
    {
        bool disposed;

        internal TempFileBuffer()
        {
            FilePath = Path.GetTempFileName();
            Stream = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }

        internal string FilePath { get; }

        internal Stream Stream { get; }

        internal long Length => Stream.Length;

        internal void Rewind()
        {
            Stream.Flush();
            Stream.Seek(0, SeekOrigin.Begin);
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            Stream.Dispose();

            try
            {
                File.Delete(FilePath);
            }
            catch
            {
            }
        }
    }
}
