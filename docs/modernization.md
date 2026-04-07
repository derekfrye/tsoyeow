# Modernization

## Goal

Modernize the codebase incrementally without losing behavioral confidence in the Excel export pipeline.

This repo is small enough to improve materially, but old enough that doing everything at once would create avoidable churn.

## Recommended Order

### 1. Establish repo-level guardrails

Start by adding shared tooling and policy rather than editing implementation code immediately.

Recommended first additions:

- root `.editorconfig`
- root `Directory.Build.props`
- analyzer setup for first-party code only

At this stage, keep warnings visible but do not fail the build on all of them yet.

### 2. Exclude noise from the first pass

Do not start with:

- `DotNetZip/`
- WinForms `*.Designer.cs`
- resource/settings generated files

That keeps the initial signal focused on real code you own.

### 3. Modernize `ExcelXmlWriter` first

This is the highest-leverage project because both the app and tests depend on it.

Current dependency state to keep in mind:

- `ExcelXmlWriter` has already been converted to an SDK-style `net10.0` project
- it no longer uses `Ionic.Zip` for XLSX packaging and now relies on `System.IO.Compression`
- it still depends on `Microsoft.Data.SqlClient`, which is a normal modern dependency for direct SQL Server access
- it still depends on `System.Configuration.ConfigurationManager`, which is a compatibility dependency tied to legacy settings usage inside the library
- the repo still contains a vendored `DotNetZip/` tree, but `ExcelXmlWriter` itself is no longer wired to it

Recommended modernization scope for the next library pass:

- complete the nullability cleanup in `ExcelXmlWriter` rather than suppressing warnings piecemeal
- normalize the event and API surface so it reads like a modern .NET library
- keep `QueryReader` implementing `IDataReader`, but make it delegate correctly to the wrapped reader instead of leaving members unimplemented
- replace the raw progress thread with task-based primitives from modern .NET
- make XML/file parsing and string formatting culture-stable and security-conscious

Priority files for review after analyzers are enabled:

- `ExcelXmlWriter/Workbook/Workbook.cs`
- `ExcelXmlWriter/QueryReader.cs`
- `ExcelXmlWriter/Xlsx/XlsxWorksheet.cs`
- `ExcelXmlWriter/Xlsx/XlsxParts.cs`

Reasons:

- they are central to core behavior
- they contain older threading and synchronization patterns
- they are likely to produce meaningful analyzer findings rather than cosmetic ones

Detailed work plan for that pass:

#### 3a. Nullability and event/API cleanup

Primary targets:

- `ExcelXmlWriter/Workbook/Workbook.cs`
- `ExcelXmlWriter/Workbook/WorkbookEvents.cs`
- `ExcelXmlWriter/QueryReader.cs`
- `ExcelXmlWriter/Xlsx/XlsxWorksheet.cs`
- `ExcelXmlWriter/Xlsx/XlsxAppXml.cs`
- `ExcelXmlWriter/Xlsx/XlsxWorkbookMetadata.cs`
- `ExcelXmlWriter/Xlsx/XlsxSharedStringsXml.cs`

Objectives:

- resolve the current `CS860x`, `CS8618`, and `CS8625` warnings by making nullability explicit
- mark events nullable where that reflects real usage
- initialize or narrow fields that are only valid after construction or after a specific lifecycle step
- replace chains of XML dereferences with guarded lookups so the code expresses the expected document shape clearly
- rename event payload types to end in `EventArgs` so the public surface matches .NET conventions

Guardrails:

- prefer making lifecycle assumptions explicit over using null-forgiving operators
- avoid changing externally visible behavior unless the change is clearly a correctness fix
- keep public API churn intentional and documented, especially if event type names change

#### 3b. `QueryReader` delegation cleanup

`QueryReader` should remain the adapter around the active reader, but it should stop being a half-implemented `IDataReader`.

Objectives:

- delegate all `IDataReader` and `IDataRecord` members to the wrapped `IDataReader`
- preserve the existing row-consumption bookkeeping where it is behaviorally important
- make the wrapped-reader lifecycle obvious for both SQL-backed and file-backed execution
- remove `NotImplementedException` members from normal runtime paths

Design notes:

- keep the wrapper focused on the repo-specific result-set and counting behavior
- where delegation requires an open reader, fail with a clear state exception rather than relying on null dereferences
- if a member is truly unsupported, document why and reconsider whether implementing `IDataReader` is still the correct abstraction

#### 3c. Replace raw-thread progress reporting

Current progress reporting uses a manually managed background `Thread`, shared mutable state, and `Thread.Sleep`.

Objectives:

- move progress emission in `Workbook` to `Task`, `CancellationToken`, and `PeriodicTimer`
- replace ad hoc cancellation flags with a clear cancellation source owned by the workbook lifecycle
- ensure shutdown paths cancel and await progress work consistently
- keep the current user-visible event cadence unless there is a reason to change it

Guardrails:

- do not introduce concurrent reads of the underlying reader beyond what the current design already assumes
- make disposal and cancellation idempotent
- keep the implementation simple enough that tests can validate completion and cancellation behavior

#### 3d. Safer XML loading for file-backed reads

The file-backed query path currently uses `DataSet.ReadXml(query)` directly.

Objectives:

- switch to an `XmlReader`-based load path
- disable DTD processing and avoid permissive resolver behavior
- preserve the existing XML-to-`DataSet` behavior expected by tests and the GUI flow

Guardrails:

- validate the new load path against the existing file-backed fixtures before and after the change
- keep locale-sensitive dataset behavior explicit where it matters

#### 3e. Culture-stable formatting and string comparisons

Spreadsheet generation should not vary by machine locale.

Objectives:

- replace culture-sensitive `ToString()` calls in serialization paths with explicit invariant formatting where appropriate
- standardize string comparisons to `Ordinal` or `OrdinalIgnoreCase` unless there is a documented business reason to use culture-aware comparison
- review workbook naming, relationship IDs, numeric serialization, and shared-string comparisons first

Guardrails:

- treat output-file stability as more important than matching legacy locale quirks
- be explicit about the few places where culture-aware behavior is genuinely required

#### 3f. Constants and duplicated value-writing logic

There are still a few embedded tuning values and duplicated write paths in the XLSX writer.

Objectives:

- replace magic numbers with named constants where the value expresses a deliberate limit or formatting rule
- extract the overpunch precision value into a named constant instead of repeating `100`
- extract the shared-string cache limit into a named constant instead of repeating a bare integer
- consolidate header-cell and row-cell value writing so type resolution and overpunch handling flow through one shared path

Design notes:

- prefer a small set of internal constants near the classes that own the behavior
- keep the constants named after the business intent rather than the literal value
- avoid duplicating the same `ResolveDataType` and `applyOverPunch` decisions across header and row-writing code

Guardrails:

- do not change output behavior just to make the code look tidier
- keep the shared helper small enough that the write path stays readable

#### 3g. Shared-string cache under memory pressure

The current shared-string cache uses hash-based lookup to limit memory, but its collision behavior is too weak.

Objectives:

- replace the hash-keyed cache with a capped `Dictionary<string, long>` that uses `StringComparer.Ordinal`
- keep the cache explicitly bounded because large-result exports can create real memory pressure
- once the cache cap is reached, continue writing shared strings without caching newly encountered values
- accept the resulting duplicate shared-string entries beyond the cap as a deliberate tradeoff for bounded memory usage and fast writes

Design notes:

- start with a named cap such as `100_000` cached strings and treat it as a tuning value
- prefer a straightforward ordinal string dictionary over bespoke hash-bucket logic
- document that the cache is an optimization layer, not a canonical deduplication guarantee once the cap is hit

Guardrails:

- optimize for predictable memory use over perfect global deduplication
- do not reintroduce collision ambiguity by storing a single entry per hash bucket
- keep the write path append-friendly even after the cache stops admitting new entries

#### 3h. Schema metadata caching per result set

Some hot paths repeatedly call `GetSchemaTable()` and rescan column metadata during row processing.

Objectives:

- read schema metadata once at the start of each result set
- derive a compact per-result-set structure from that schema and reuse it through sheet initialization and key-column lookup
- avoid repeated `GetSchemaTable().Rows.Cast<DataRow>()` scans in row-writing logic

Design notes:

- cache the raw schema table once per result set where needed for sheet creation
- also derive a faster lookup structure for column presence checks, such as a `HashSet<string>` with `StringComparer.OrdinalIgnoreCase`
- pass the cached schema metadata into the places that currently pull it back from the reader repeatedly

Guardrails:

- keep the cache scoped to a single result set so schema changes across result sets remain correct
- do not assume all result sets have the same columns
- verify the cached structure matches the exact metadata needed by worksheet creation and duplicate-key checks

#### 3i. Temp-file and stream lifecycle cleanup

The XLSX writer still uses manual temp-file creation, manual stream rewinds, and repeated best-effort delete logic.

Objectives:

- keep disk-backed buffering as the preferred strategy for large-result-set exports
- avoid switching core XLSX assembly work to in-memory buffering by default
- introduce one internal abstraction that owns the temp file path, backing stream, rewind behavior, and cleanup
- remove duplicated `Path.GetTempFileName()` plus `FileStream` setup from worksheet and shared-string code
- centralize best-effort delete behavior so cleanup policy lives in one place

Design notes:

- the abstraction should make stream ownership unambiguous
- callers should not need to coordinate a separate file path and stream manually
- cleanup should be idempotent and safe even when close and dispose paths are both exercised
- if cleanup failures remain best-effort, contain that policy in one helper rather than scattering empty `catch` blocks

Guardrails:

- preserve the current large-output behavior that avoids holding full workbook content in memory
- do not introduce a `MemoryStream`-first design for the main export path unless benchmarks or tests justify it
- keep rewind-before-zip-copy behavior explicit and testable

Suggested execution order inside `ExcelXmlWriter`:

1. Nullability fixes and event/API cleanup.
2. `QueryReader` delegation cleanup.
3. Task-based progress and cancellation refactor in `Workbook`.
4. `XmlReader`-based file-backed load path.
5. Invariant formatting and ordinal comparison cleanup.
6. Constants cleanup, shared-string cache redesign, and schema metadata caching.
7. Temp-file and disk-backed stream lifecycle cleanup.
8. Rebuild and rerun library and regression tests before taking on lower-value analyzer noise.

### 4. Use tests to hold the line

After analyzer setup and project-file work, lean on:

- `ExcelXmlWriterNTest/WorkbookTest.cs`
- `ExcelXmlWriterNTest/ConfigManipulatorTest.cs`

The tests are also legacy, but they still provide useful regression coverage while the library evolves.

### 5. Tackle the WinForms app second

Once the library is under better tooling, move to:

- `ExcelXmlQueryResults/FormEntrance.cs`
- `ExcelXmlQueryResults/FormOptions.cs`
- `ExcelXmlQueryResults/ConfigManipulator.cs`

This phase should focus on code quality and maintainability, not UI redesign.

Configuration work should be split into two tracks:

- library track: remove `System.Configuration.ConfigurationManager` from `ExcelXmlWriter` by replacing `Settings.Default` usage with embedded resources or explicit inputs
- app track: migrate the WinForms app's `App.config` read/write flow to a more modern configuration approach if that still provides enough value

Those are related, but they are not the same task and should not be treated as a single migration.

### 6. Reassess `DotNetZip`

Only after the first-party code is cleaner should you decide whether to:

- keep the vendored source tree for historical/reference purposes
- replace any remaining vendored/legacy usage with a maintained packaging approach
- remove the local `DotNetZip` source tree if it is no longer needed

## Practical First Milestone

A strong first modernization milestone would be:

1. Add docs and repo-level policy files.
2. Enable analyzers for first-party code.
3. Convert `ExcelXmlWriter` to a modern project format.
4. Remove the library's remaining legacy configuration dependency if feasible without broad app changes.
5. Complete the nullability and API cleanup pass in the core library.
6. Refactor `QueryReader`, progress reporting, XML loading, and culture-sensitive formatting in the library.
7. Confirm tests still run.

That sequence improves maintainability quickly without forcing a full framework migration on day one.

## Things To Preserve

During modernization, keep these constraints in mind:

- existing output behavior matters more than stylistic purity
- large-result-set handling is core to the repo's value
- historical debug configurations may encode useful troubleshooting modes
- generated and vendored files should not dominate review noise

## Longer-Term Options

After the initial cleanup, the larger decisions are:

- migrate to SDK-style projects throughout the repo
- move from `.NET Framework 4.0` to a newer target where practical
- migrate app-layer configuration separately from library internals
- replace manual threading with more modern task-based patterns
- improve test ergonomics and reduce fixture brittleness
- separate the core export engine more cleanly from the WinForms front end
