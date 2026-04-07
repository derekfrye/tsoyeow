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

Priority files for review after analyzers are enabled:

- `ExcelXmlWriter/Workbook/Workbook.cs`
- `ExcelXmlWriter/QueryReader.cs`
- `ExcelXmlWriter/Xlsx/XlsxWorksheet.cs`
- `ExcelXmlWriter/Xlsx/XlsxParts.cs`

Reasons:

- they are central to core behavior
- they contain older threading and synchronization patterns
- they are likely to produce meaningful analyzer findings rather than cosmetic ones

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

### 6. Reassess `DotNetZip`

Only after the first-party code is cleaner should you decide whether to:

- keep the existing dependency as-is
- replace vendored/legacy usage with a maintained packaging approach
- remove the local `DotNetZip` source tree if it is no longer needed

## Practical First Milestone

A strong first modernization milestone would be:

1. Add docs and repo-level policy files.
2. Enable analyzers for first-party code.
3. Convert `ExcelXmlWriter` to a modern project format.
4. Fix the first wave of analyzer findings in the core library.
5. Confirm tests still run.

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
- replace manual threading with more modern task-based patterns
- improve test ergonomics and reduce fixture brittleness
- separate the core export engine more cleanly from the WinForms front end
