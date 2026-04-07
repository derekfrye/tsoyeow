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

Current implementation status for the library:

- done: the library has been restructured around a new internal export session and a thinner `Workbook` compatibility facade
- done: `QueryReader` now delegates `IDataReader` members instead of leaving most of the interface unimplemented
- done: progress reporting now uses task-based primitives rather than a raw background `Thread`
- done: file-backed dataset loading now goes through `XmlReader`
- done: schema metadata is cached once per result set and reused in the write path
- done: the shared-string cache now uses a capped `Dictionary<string, long>` with ordinal comparison
- done: repeated XLSX value-writing logic has been consolidated and some magic numbers have been extracted to named constants
- done: public event payloads were renamed to modern `*EventArgs` names and the GUI source was updated to use them
- done: the main nullability and culture/formatting fixes in the refactored library write path have been applied
- not done: `System.Configuration.ConfigurationManager` is still used by the library
- not done: a true Excel-compatible XML backend has not yet been implemented
- not done: the disk-backed temp-buffer abstraction described below has not yet been introduced
- not done: lower-priority analyzer/style cleanup is still pending
- not done: full regression-test verification for the refactor has not been completed

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

Status:

- partially done: event payload types were renamed to `*EventArgs`, `Workbook` event declarations were updated, and the GUI source was updated to match
- partially done: the main XML helper nullability hazards in the XLSX path were replaced with explicit required-element checks
- not done: some lower-value nullability/analyzer cleanup still remains in the library

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

Status:

- done: `QueryReader` now acts as a real delegating adapter over the active reader for both SQL-backed and file-backed execution

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

Status:

- done: progress emission was moved to task-based primitives in the new export session layer

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

Status:

- done: file-backed dataset loading now uses `XmlReader` with restrictive settings

Guardrails:

- validate the new load path against the existing file-backed fixtures before and after the change
- keep locale-sensitive dataset behavior explicit where it matters

#### 3e. Culture-stable formatting and string comparisons

Spreadsheet generation should not vary by machine locale.

Objectives:

- replace culture-sensitive `ToString()` calls in serialization paths with explicit invariant formatting where appropriate
- standardize string comparisons to `Ordinal` or `OrdinalIgnoreCase` unless there is a documented business reason to use culture-aware comparison
- review workbook naming, relationship IDs, numeric serialization, and shared-string comparisons first

Status:

- partially done: the refactored write path now uses invariant formatting and ordinal comparison in the most important serialization and cache paths
- not done: lower-priority culture/analyzer cleanup still remains in some library helpers

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

Status:

- partially done: overpunch precision and shared-string cache limits are now named constants in the refactored XLSX path
- done: header and row value writing now flow through one shared worksheet write path

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

Status:

- done: the shared-string cache now uses a capped ordinal `Dictionary<string, long>`

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

### 4. Expand modern test coverage

The new SDK-style `xUnit` suite covers the current file-backed export path and core XLSX internals, but a few high-value behaviors still need explicit tests.

Recommended next additions:

- add SQL-backed integration tests for `QueryReader` and `ExcelExporter`, ideally using a disposable SQL Server fixture rather than machine-local config
- add cancellation and progress-loop tests for `WorkbookExportSession`, including cancellation after `Open()` and during multi-workbook export
- add worksheet-boundary tests for `DupeKeysToDelayStartingNewWorksheet`, especially the case where the row limit is reached inside a duplicate-key group
- add explicit package-metadata tests for `XlsxWorkbookMetadata`, `XlsxAppXml`, `ContentTypes`, and `Relationships` so workbook manifests are validated directly instead of only through broad XLSX smoke tests
- add malformed and hostile XML fixture tests for the file-backed path, including prohibited DTD input and schema-shape failures

Status:

- done: result-set schema metadata is now cached once and reused through the write path

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

Status:

- partially done: the refactor kept disk-backed buffering and reduced some duplication in the XLSX path
- not done: there is still no dedicated temp-buffer abstraction with centralized lifecycle ownership

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

### 5. Tackle the desktop app second

Once the library is under better tooling, move to:

- `ExcelXmlQueryResults/FormEntrance.cs`
- `ExcelXmlQueryResults/FormOptions.cs`
- `ExcelXmlQueryResults/ConfigManipulator.cs`

The current GUI is a legacy `.NET Framework 4.0` WinForms application. If the goal is a desktop GUI that runs on both Windows and Linux, the right modernization path is to replace that front end with a cross-platform desktop shell rather than carrying WinForms forward.

Recommended direction:

- keep `ExcelXmlWriter` as the export engine and avoid moving workbook logic into the new UI
- build the replacement desktop app on Avalonia rather than trying to make WinForms cross-platform
- treat the GUI migration as an architectural cleanup opportunity, not just a visual redesign

The current UI behavior is modest and portable:

- query text entry
- destination file picking
- batch query-file selection
- options editing for connection and workbook settings
- progress and error reporting while exports run

That behavior should be reproduced, but the new app should not preserve the current WinForms structure.

#### 5a. Extract application services before or during the GUI port

Before binding a new UI to the existing code, move non-UI logic out of forms and into plain modern classes.

Primary extraction targets:

- export workflow orchestration currently embedded in `FormEntrance`
- settings load/save logic currently embedded in `FormOptions` and `ConfigManipulator`
- connection-string building
- output-path validation and file naming
- progress/state projection from workbook events into UI-friendly state objects

Target architecture:

- `ExcelXmlWriter`:
  owns export behavior only
- application/service layer:
  owns settings, workflow coordination, validation, and progress state
- Avalonia UI layer:
  owns view models, bindings, dialogs, and presentation only

This keeps the Avalonia app from becoming a straight port of WinForms event handlers.

#### 5b. Replace legacy config mutation with modern settings storage

The current app edits `App.config` directly with `XDocument`. That is fragile, tightly coupled to legacy executable config behavior, and a poor fit for a cross-platform desktop app.

Recommended replacement:

- use a typed settings model
- persist settings in a JSON file under the user's application data directory
- add one-time migration logic that can import the legacy config values on first run if preserving existing user settings matters

Configuration work should be split into two tracks:

- library track: remove `System.Configuration.ConfigurationManager` from `ExcelXmlWriter` by replacing `Settings.Default` usage with embedded resources or explicit inputs
- app track: replace the WinForms app's `App.config` read/write flow with a modern desktop settings approach

Those are related, but they are not the same task and should not be treated as a single migration.

#### 5c. Clean up app-layer threading and progress

Do not port the current raw-thread and `Invoke` patterns into the new UI.

Recommended changes:

- replace raw `Thread` usage with `Task` and `async`/`await`
- use `CancellationToken` for cancellation and shutdown
- expose progress through a narrow state/progress abstraction rather than direct form callbacks
- make the UI consume state changes instead of controlling worker lifetime directly

This work belongs mostly in the shared application/service layer so both tests and the new UI can exercise it.

#### 5d. Normalize the public engine surface before binding a new UI to it

The current workbook API still exposes legacy naming and lifecycle patterns. If a new UI is going to depend on it long-term, clean up that surface first instead of baking the old API into a new application shell.

Useful targets:

- rename legacy types such as `WorkBookParams` and `WorkBookStatus`
- make lifecycle expectations explicit around query execution, export continuation, and cleanup
- reduce accidental public surface area
- keep backend-specific implementation details out of the UI-facing contract

#### 5e. Resolve misleading or obsolete UI-facing options

The current app still exposes an `Xml` backend mode even though the implementation path is effectively the XLSX writer.

Resolve this before or during the GUI migration:

- either implement a true Excel-compatible XML backend
- or remove the XML backend option from both the app and public API

Do not carry forward options that are not truly supported.

#### 5f. Remove legacy-only dependencies and project structure while the UI is in motion

The GUI migration is a good time to stop carrying forward dependencies and project patterns that only exist to support the old WinForms application.

Good candidates:

- remove `System.Configuration` from the app layer
- migrate the remaining legacy projects from old project format to SDK-style projects
- replace `packages.config` usage with `PackageReference`
- standardize on `Microsoft.Data.SqlClient` rather than splitting provider usage across old and new SQL client packages
- retire WinForms designer-generated assumptions instead of recreating them in Avalonia

#### 5g. Reassess historical debug modes and one-off code paths

The repo contains historical debug configurations and conditional compilation paths such as file-driven and no-thread execution modes. Preserve them long enough to understand what behavior they protect, but do not assume they all deserve to survive the migration unchanged.

During the port:

- identify which debug modes still represent useful troubleshooting workflows
- replace the useful ones with explicit modern switches or test coverage where possible
- remove dead debug-only path walking and form-specific scaffolding once equivalent coverage exists

### 6. Reassess `DotNetZip` and other leftover legacy dependencies

Only after the first-party code is cleaner should you decide whether to:

- keep the vendored source tree for historical/reference purposes
- replace any remaining vendored/legacy usage with a maintained packaging approach
- remove the local `DotNetZip` source tree if it is no longer needed

Related cleanup opportunities:

- remove `Ionic.Zip` usage from tests if `System.IO.Compression` is sufficient
- remove test or utility dependencies that only existed to support the legacy packaging path
- keep third-party dependency removal separate from behavior changes when possible

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
