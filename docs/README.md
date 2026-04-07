# Purpose And Quick Start

This repository contains a legacy .NET solution for exporting SQL Server query results into Excel-compatible files, primarily `.xlsx` workbooks that modern Excel can open.

The codebase is organized around three first-party projects:

- `ExcelXmlWriter`: reusable library that reads SQL results and writes workbook output
- `ExcelXmlQueryResults`: WinForms desktop app that provides the user-facing workflow
- `ExcelXmlWriterNTest`: NUnit-based test project with sample resources

There is also a vendored copy of `DotNetZip`, which should generally be treated as third-party code rather than a place to start routine cleanup.

## Quick Start

This is an old-style Visual Studio solution targeting `.NET Framework 4.0`, not a modern SDK-style multi-targeted repo.

The main solution file is:

- `Tsoyeow.sln`

The most likely entry points are:

- `ExcelXmlQueryResults` for running the desktop application
- `ExcelXmlWriterNTest` for running tests and inspecting behavioral expectations

## Build Notes

The repo currently uses:

- legacy `*.csproj` project files
- `packages.config` for NuGet dependencies
- custom solution configurations such as `DebugFromFile`, `DebugNoThread`, and `DebugFromFileNoThread`

If you open the solution in Visual Studio, start with the default `Debug|Any CPU` configuration unless you specifically need one of the historical debug modes.

## Project Docs

- [architecture.md](./architecture.md): codebase layout and project boundaries
- [development.md](./development.md): practical notes for building, testing, and editing safely
- [modernization.md](./modernization.md): recommended modernization order and rationale
