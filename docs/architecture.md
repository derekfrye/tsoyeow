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
