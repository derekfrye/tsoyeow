using System.Globalization;

namespace ExcelXmlWriter.Workbook
{
    internal sealed partial class WorkbookExportSession
    {
        void InitializeSheet(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema)
        {
            backend.CreateSheet(
                queryReader.CurrentResult,
                tracking.SheetSubCount,
                RetrieveSheetName(queryReader.CurrentResult, tracking.SheetSubCount),
                schema);
            tracking.WorksheetOpen = true;
        }

        void DetermineIfRowDependsOnPreviousRow(WorkbookTracking tracking, ResultSetSchema schema)
        {
            tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey = ReadDuplicateKeyValues(schema);
            if (tracking.RowCount == 1)
            {
                tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey = tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;
            }

            if (tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey is not null
                && tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey is not null)
            {
                tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = tracking.PreviousAndCurrentRowKeyColumns
                    .CurrentRowDupKey
                    .Zip(tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey, (current, previous) => string.Equals(current, previous, StringComparison.Ordinal))
                    .Any(matches => !matches);
            }
            else
            {
                tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = true;
            }

            tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey = tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;
        }

        string[]? ReadDuplicateKeyValues(ResultSetSchema schema)
        {
            if (options.DupeKeysToDelayStartingNewWorksheet.Length == 0)
            {
                return null;
            }

            var values = new string[options.DupeKeysToDelayStartingNewWorksheet.Length];
            for (int index = 0; index < values.Length; index++)
            {
                var columnName = options.DupeKeysToDelayStartingNewWorksheet[index];
                if (schema.ContainsColumn(columnName))
                {
                    values[index] = queryReader[columnName]?.ToString() ?? string.Empty;
                }
            }

            return values;
        }

        string RetrieveSheetName(int sheetCount, int sheetSubCount)
        {
            if (options.ResultNames.TryGetValue(sheetCount, out var sheetName))
            {
                return sheetName + "_" + sheetSubCount.ToString(CultureInfo.InvariantCulture);
            }

            return "Sheet" + sheetCount.ToString(CultureInfo.InvariantCulture) + "_" + sheetSubCount.ToString(CultureInfo.InvariantCulture);
        }
    }
}
