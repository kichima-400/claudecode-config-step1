---
name: vba-naming-and-data
description: Expanded rules for naming, variable typing, enums, data structures, and clarity.
---
# VBA Naming, Types & Data Management — Expanded Edition

*Clear, consistent naming and strong typing are the backbone of reliable VBA applications.*

## Purpose

Improved correctness, maintainability, and readability by applying unified naming patterns and predictable data structures.

## Naming Best Practices

### Procedures

- Use **verbs** for Subs: `LoadData`, `ExportReport`, `ValidateInputs`.
- Use **value-returning names** for Functions: `GetLastRow`, `IsCustomerValid`.

### Variables

Use names that express: *what*, not *how*.
```vb
Dim reportDate As Date
Dim customerTotal As Currency
Dim pendingInvoices As Collection
```
Avoid vague terms:
```vb
Dim tmp, x, data1 ' BAD
```

### Constants

```vb
Private Const CONFIG_SHEET As String = "Config"
Private Const HEADER_ROW As Long = 1
```

## Type Best Practices

### Good defaults

- `Long` for counters and row indexes.
- `Date` for dates.
- `Currency` for money.
- `Double` for scientific/decimal work.
- `Boolean` for flags.

### Enums

Replace magic numbers:
```vb
Public Enum ReportType
    rptSummary = 1
    rptDetail = 2
End Enum
```

### User-Defined Types

Lightweight structured groups:
```vb
Private Type TRangeBounds
    FirstRow As Long
    LastRow As Long
End Type
```

### Collections & Dictionaries

Use dictionaries for keyed access.
```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

## Do / Don't

### Do

- Use descriptive variable names.
- Use Enums instead of numbers.
- Use concrete types to avoid runtime errors.

### Don't

- Use `Variant` when you know the type.
- Mix naming conventions arbitrarily.
- Store "Y/N" as strings instead of Boolean.

## Acceptance Checklist

- [ ] No meaningless variable names.
- [ ] All numeric codes replaced by Enums.
- [ ] Data grouped logically using `Type` or classes.
