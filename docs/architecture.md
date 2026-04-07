# Architecture

## High-Level Overview

The repository is a small multi-project .NET solution built around one central library and one desktop application.

## Projects

### `ExcelXmlWriter`

This is the core library. Its responsibilities appear to include:

- connecting to SQL Server and reading result sets
- tracking workbook state and row progress
- writing spreadsheet parts for `.xlsx` output
- handling multiple result sets and workbook/tab emission

Important areas:

- `Workbook/`: workbook lifecycle, worker coordination, events, and status tracking
- `Xlsx/`: low-level workbook-part generation for `.xlsx` files
- `QueryReader.cs`: SQL data-reader orchestration
- `StaticFunctions.cs`: parsing and conversion helpers

Known architecture TODO:

- add or restore a true Excel-compatible XML backend instead of treating the `Xml` backend mode as an alias for `.xlsx` generation
- redesign the library around a narrower public facade and clearer internal component boundaries for long-term maintainability

Long-term redesign direction for this library:

- keep `Workbook` as the main public entry point only if it becomes a thin facade rather than the current orchestration-heavy type
- reduce the public surface to the types that are intentionally supported by the repo and future callers
- treat backend writers, query adapters, temp-buffer handling, and package-part generation as internal implementation details
- optimize for predictable large-result-set behavior and bounded memory use over convenience abstractions that require full in-memory workbook models

Target public API shape:

- one primary public workflow type for running an export
- one public options/request type describing query source, connection/input mode, backend choice, and export limits
- a small set of public status and progress types with conventional .NET naming
- explicit support for both `.xlsx` and Excel-compatible XML only when both backends are truly implemented

Public API cleanup goals:

- rename legacy API shapes such as `WorkBookParams` and `WorkBookStatus` to conventional modern names
- rename event payload types to `*EventArgs` and expose PascalCase properties
- remove accidental public implementation types such as `XlsxParts` unless there is a deliberate reason to support them directly
- make lifecycle expectations explicit so callers do not need to memorize fragile call ordering across `RunQuery`, `WriteQueryResults`, and `QueryClose`

Preferred responsibility split:

- `Workbook` or successor facade:
  coordinates a single export operation and exposes public events or progress reporting
- query source / reader adapter:
  owns SQL-backed vs file-backed reader creation and exposes a consistent result-set reader abstraction
- export session / coordinator:
  owns workbook splitting, worksheet splitting, duplicate-key boundary rules, and high-level write flow
- backend implementation:
  owns workbook-format-specific writing for `.xlsx` or Excel XML
- temp-buffer abstraction:
  owns disk-backed buffering, stream rewinds, and cleanup semantics
- value conversion layer:
  owns Excel data typing, overpunch handling, invariant formatting, and shared reusable write decisions
- schema metadata cache:
  owns per-result-set schema lookup structures instead of rereading schema from the live reader in hot paths

`Workbook` redesign goals:

- keep it as the user-facing facade if that remains the clearest API, but move orchestration internals out of it
- avoid mixing query execution, result-set iteration, split policy, backend selection, and progress-loop lifecycle in one class
- make cancellation, completion, and disposal semantics explicit and testable
- support both one-shot export flows and, only if still needed, incremental continuation flows for oversized outputs

`QueryReader` redesign goals:

- keep a wrapper/adapter only if it provides real value as the single abstraction over SQL-backed and file-backed readers
- if retained, make it a focused delegating adapter over the active `IDataReader`
- remove the current half-implemented interface shape and ensure all delegated members behave consistently
- keep row counting, result-set tracking, and source-specific setup in this layer rather than leaking those concerns into `Workbook`
- split it further only if the adapter still carries too many unrelated responsibilities after delegation cleanup

Backend redesign goals:

- support a true Excel-compatible XML backend alongside the `.xlsx` backend
- keep backend selection behind an internal factory or equivalent composition point
- require each backend to implement the same narrow contract for sheet creation, row writing, file-size tracking, and close/finalize behavior
- do not let one backend type leak into the public API as the de facto supported model

XLSX writer redesign goals:

- keep the low-level, hand-built, streaming-oriented approach if that remains the best fit for large exports
- stop concentrating XML structure assembly, value conversion, packaging, temp-buffer handling, and shared-string policy into a few broad classes
- isolate the current handwritten XML generation behind narrower helpers such as worksheet writer, shared-string manager, package assembler, and template/resource provider
- prefer explicit streaming helpers over introducing a full in-memory workbook object model

Performance and memory goals:

- preserve disk-backed buffering as the normal path for large exports
- keep memory usage bounded and predictable, even if that means accepting partial deduplication of shared strings after a cache cap
- cache per-result-set schema metadata once instead of repeatedly traversing schema tables during hot row-writing paths
- validate redesign work against the repo's large-result-set scenarios rather than only against small functional tests

Testing implications of the redesign:

- expand regression coverage around lifecycle transitions, oversize workbook continuation, duplicate-key sheet splitting, and file-backed dataset input
- test both backend implementations through the same higher-level contract where possible
- add focused tests around progress/cancellation semantics and shared-string caching behavior
- keep golden-file or fixture-based verification for output compatibility where practical

This is the best place to start when changing behavior because both the app and tests depend on it.

### `ExcelXmlQueryResults`

This is the WinForms front end. Its responsibilities appear to include:

- collecting connection and query options
- kicking off export work
- handling application configuration
- presenting progress and file-output options

Important areas:

- `FormEntrance.cs`: main workflow and threaded execution handoff
- `FormOptions.cs`: application options and configuration UI
- `ConfigManipulator.cs`: config persistence and related utility logic

Designer files under this project are generated artifacts and should usually not be the first place to make structural changes.

### `ExcelXmlWriterNTest`

This is the test project. It contains:

- workbook-related tests
- config manipulation tests
- sample SQL and workbook fixture data in `Resources/`

The project references both first-party projects, so it can act as a basic regression net during modernization.

### `DotNetZip`

This directory contains vendored zip/compression source code. It targets much older framework versions and has its own project files and code-analysis settings.

For routine repo cleanup, treat this as third-party code:

- do not mix it into initial lint cleanup
- do not reformat it casually
- only touch it if there is a concrete packaging or dependency reason

## Dependency Direction

The dependency flow is straightforward:

- `ExcelXmlQueryResults` -> `ExcelXmlWriter`
- `ExcelXmlWriterNTest` -> `ExcelXmlWriter`
- `ExcelXmlWriterNTest` -> `ExcelXmlQueryResults`

`DotNetZip` is consumed via the `Ionic.Zip` package reference path in legacy project files rather than as an active first-party project dependency in the main solution.

## Design Characteristics

A few architectural traits stand out immediately:

- the code predates async/await and relies on manual `Thread` usage
- some data and workbook workflows are coordinated with explicit `lock` usage
- the output layer includes hand-built XML and package-part generation
- a meaningful amount of WinForms designer code is present

That combination means modernization should start with project/build scaffolding and core-library cleanup before any UI redesign effort.
