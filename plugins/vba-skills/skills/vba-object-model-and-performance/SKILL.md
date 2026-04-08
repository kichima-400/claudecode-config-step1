---
name: vba-objectmodel-performance
description: Safely use the Office object model; perform fast, scalable operations using arrays and batching.
---
# Safe Object Model Usage & Performance — Expanded Edition

*Avoid UI dependence and dramatically improve performance with array processing and object-model best practices.*

## Purpose

Improve reliability and speed by eliminating fragile UI-linked code and minimizing expensive cross-boundary calls between VBA and Office.

## Object Model Safety

### Never Use `Select`, `Activate`

```vb
' BAD
aRange.Select
Selection.Value = 5

' GOOD
ws.Range("A1").Value = 5
```

### Fully Qualify Everything

```vb
ThisWorkbook.Worksheets("Report").Range("A1").Value = "OK"
```

### Use Stable Anchors

- Named ranges
- Tables (`ListObjects`)
- Constants for key locations

## Worksheet I/O Best Practices

### Use Array Round-Trips

```vb
Dim arr As Variant
arr = ws.Range("A1:A100000").Value2

' Modify in memory
For i = 1 To UBound(arr)
    arr(i, 1) = arr(i, 1) * 2
Next i

ws.Range("A1:A100000").Value2 = arr
```

### Cache References

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Data")
```

### Turn Off Expensive Features

```vb
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
```
Restore afterwards.

## Performance Do / Don't

### Do

- Use arrays for all large loops.
- Disable/unset features inside try/finally-style blocks.
- Batch formatting operations.

### Don't

- Process ranges cell-by-cell inside loops.
- Repeatedly call `WorksheetFunction` inside loops.
- Use `ReDim Preserve` inside loops.

## Anti-Patterns

BAD: Code that runs only when the correct sheet is selected.
BAD: Large number of cell updates performed one-by-one.
BAD: Repeated workbook/worksheet lookups in inner loops.

## Acceptance Checklist

- [ ] All automation works with no UI context.
- [ ] Arrays used for heavy lifting.
- [ ] Performance settings restored safely.
