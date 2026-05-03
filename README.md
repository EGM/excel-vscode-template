# excel-vscode-template

A modern VBA development workflow using:

- VSCode as the editor
- Git as the source of truth
- Rubberduck VBA for testing
- PowerShell scripts for import/export
- A configurable workbook path stored in `config.json`

## Quick Start

1. Clone this repo
2. Run:

```ps
./scripts/configure.ps1
```

3. Open the folder in VSCode:

```ps
code .
```

4. Use VSCode tasks:
- **Ctrl+Shift+B** → Import code into Excel
- **Ctrl+Shift+E** → Export code from Excel
- **Ctrl+Shift+T** → Run Rubberduck tests

Your workbook stays clean.  
Your source code stays in Git.  
Your tests run headlessly.

## Folder Layout

```ps
src/        → VBA source modules
tests/      → Rubberduck test modules
scripts/    → PowerShell automation
.vscode/    → VSCode tasks
config.json → Workbook path + settings
```

## First-Time Setup

Run:

```ps
./scripts/configure.ps1
```

This will:

- Ask for a workbook name
- Create it if missing
- Write the path into `config.json`

All other scripts read from that file.
