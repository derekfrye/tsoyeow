using ExcelXmlWriter.Xlsx;

namespace ExcelXmlWriter.Workbook
{
    internal sealed partial class WorkbookExportSession : IDisposable
    {
        readonly WorkBookParams options;
        readonly QueryReader queryReader;
        readonly CancellationToken cancellationToken;

        CancellationTokenSource? progressCancellationSource;
        Task? progressTask;
        bool queryOpened;
        bool closed;
        object?[]? bufferedRow;

        internal WorkbookExportSession(WorkBookParams options, CancellationToken cancellationToken = default)
        {
            this.options = options;
            this.cancellationToken = cancellationToken;
            queryReader = new QueryReader(options.Query, options.QueryTimeout, options.FromFile, options.ConnectionString);
        }

        internal event EventHandler<EventArgs>? QueryStarted;

        internal event EventHandler<ExportProgressEventArgs>? QueryProgress;

        internal event EventHandler<SaveStartedEventArgs>? SaveStarted;

        internal event EventHandler<ExportFinishedEventArgs>? Finished;

        internal int TotalRecordsRead => queryReader.TotalRecordsRead;

        internal void Open()
        {
            cancellationToken.ThrowIfCancellationRequested();
            QueryStarted?.Invoke(this, EventArgs.Empty);
            queryReader.OpenReader();
            queryOpened = true;
            StartProgressLoop();
        }

        internal void Close()
        {
            if (closed)
            {
                return;
            }

            closed = true;
            queryReader.CloseReader();
            StopProgressLoop();
            Finished?.Invoke(this, new ExportFinishedEventArgs(queryReader.TotalRecordsRead));
        }

        static IExcelBackend CreateBackend(Stream stream)
        {
            return new XlsxParts(stream);
        }

        void StartProgressLoop()
        {
            progressCancellationSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            progressTask = Task.Run(() => EmitRowsOverTimeAsync(progressCancellationSource.Token), CancellationToken.None);
        }

        void StopProgressLoop()
        {
            if (progressCancellationSource is null)
            {
                return;
            }

            progressCancellationSource.Cancel();

            try
            {
                progressTask?.GetAwaiter().GetResult();
            }
            catch (OperationCanceledException)
            {
            }
            finally
            {
                progressCancellationSource.Dispose();
                progressCancellationSource = null;
                progressTask = null;
            }
        }

        async Task EmitRowsOverTimeAsync(CancellationToken cancellationToken)
        {
            var start = DateTime.UtcNow;
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(5));

            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                var totalRows = queryReader.TotalRecordsRead;
                var elapsed = DateTime.UtcNow - start;
                QueryProgress?.Invoke(
                    this,
                    new ExportProgressEventArgs((decimal)totalRows / Math.Max((decimal)elapsed.TotalSeconds, 0.001m), totalRows));
            }
        }

        void OnSaveStarted(string message)
        {
            SaveStarted?.Invoke(this, new SaveStartedEventArgs(message));
        }

        public void Dispose()
        {
            if (!closed)
            {
                Close();
            }

            queryReader.Dispose();
        }
    }
}
