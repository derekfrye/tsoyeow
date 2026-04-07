# Development

## Current State

This repository is a legacy `.NET Framework 4.0` solution using pre-SDK project files and `packages.config`.

Expect a development workflow closer to classic Visual Studio than to a modern `dotnet` CLI-first repo.

## Recommended Working Boundaries

When making routine changes, focus on first-party code only:

- `ExcelXmlWriter`
- `ExcelXmlQueryResults`
- `ExcelXmlWriterNTest`

Avoid broad cleanup in:

- `DotNetZip/`
- generated `*.Designer.cs` files
- auto-generated resource/settings files unless regeneration is part of the task

## Useful Project Roles

- `ExcelXmlWriter`: safest place for core behavior fixes and modernization work
- `ExcelXmlQueryResults`: app behavior and configuration UX
- `ExcelXmlWriterNTest`: regression checks and fixture-based confidence

## Historical Build Configurations

The solution contains several old custom configurations in addition to normal `Debug` and `Release`, including:

- `DebugFromFile`
- `DebugNoThread`
- `DebugFromFileNoThread`

These suggest earlier troubleshooting paths for file-driven and single-threaded scenarios. Preserve them until they are understood, even if they eventually get retired during modernization.

## Dependencies

Current notable package usage:

- `Ionic.Zip` via `packages.config`
- `NUnit` and `NUnit3TestAdapter` in the test project

Because the repo still uses package restore via `packages.config`, dependency modernization should be deliberate and separated from behavior changes where possible.

## Cleaning

There is a root PowerShell script:

- `clean.ps1`

It removes historical build output directories and also edits test config values. Treat it as a legacy maintenance script and review it before relying on it in a modern workflow.

## Safe Modernization Practices

- introduce repo-level documentation and build policy before mass cleanup
- add analyzers in a non-blocking mode first
- modernize one first-party project at a time
- keep behavioral changes separate from tooling/project-file changes
- preserve test fixtures and sample workbook resources during refactors

## Suggested Near-Term Tooling

Once modernization starts, the repo will benefit from:

- a root `.editorconfig`
- a root `Directory.Build.props`
- analyzer configuration that excludes generated and vendored code
- eventual migration from `packages.config` to `PackageReference`
- eventual migration from legacy project format to SDK-style projects
