---
name: vba-core-architecture
description: Deep guidance on professional-grade VBA architecture, module structure, and project setup.
---
# VBA Core Architecture

*Robust architectural foundations for stable, maintainable, enterprise-friendly VBA projects.*

## Purpose

Create a disciplined project structure that prevents code sprawl, reduces debugging time, improves readability, and enables multiple developers (or an LLM agent) to work consistently.

## When to Use

- Starting a new automation or add-in project.
- Refactoring legacy workbooks containing years of macros scattered across modules.
- Turning quick prototypes into production-quality automations.
- Enforcing coding standards across a team.

## Core Principles

1. **Explicit over implicit** — declare everything, qualify everything.
2. **Separation of concerns** — UI ≠ logic ≠ I/O.
3. **Thin event modules** — events route; they do not compute.
4. **Predictable structure** — all contributors know where code belongs.

## Recommended Project Structure

```
VBAProject
│
├─ modEntryPoints            ' User-facing entry macros
├─ modConfig                 ' Constants, settings
├─ modUtilities              ' Pure helpers (Option Private Module)
├─ modWorksheetIO            ' Read/write helpers
├─ modErrors                 ' Logging, error utilities
│
├─ clsApplication            ' Main orchestrator
├─ clsLogger                 ' Logging abstraction
├─ clsWorkbookGateway        ' Encapsulates workbook/range access
├─ clsDomain_*               ' Domain objects (Invoice, Customer...)
│
├─ Sheet1/SheetX             ' Thin event routers only
└─ ThisWorkbook              ' Workbook-level events only
```

## Instructions (for the LLM)

1. Insert `Option Explicit` at top of every module.
2. Mark helper modules with `Option Private Module` to hide internals.
3. For any event (Worksheet_Change, Workbook_Open), create a dedicated handler:
```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    HandleSheetChange Me, Target
End Sub
```
4. Move workflow logic to an application class:
```vb
Public Sub RunMonthlyReport()
    Dim app As New clsApplication
    app.RunMonthlyReport
End Sub
```
5. Encapsulate worksheet operations inside gateway classes instead of referencing sheets directly.

## Do / Don't

### Do

- Create layers that mirror real architecture.
- Keep code discoverable: developers know "where logic lives".
- Use one entry point per workflow.

### Don't

- Scatter logic across 20+ sheet modules.
- Store state in public global variables.
- Place complex logic directly inside event handlers.

## Anti-Patterns

BAD: God modules containing thousands of lines with mixed responsibilities.
BAD: Macros dependent on user interface (active cell, selection).
BAD: Spreadsheet logic embedded inside UI or event code.

## Acceptance Checklist

- [ ] No business logic in sheet/workbook modules.
- [ ] All modules named meaningfully.
- [ ] Helpers hidden with `Option Private Module`.
- [ ] Entry points are 1-5 lines each.
