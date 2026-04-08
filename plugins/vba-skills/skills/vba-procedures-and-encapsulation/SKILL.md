---
name: vba-procedures-and-encapsulation
description: Detailed patterns for building small, readable, encapsulated procedures.
---
# VBA Procedures, Scope & Encapsulation — Expanded Edition

*Write small, intention-revealing, testable, object-oriented VBA.*

## Purpose

Reduce hidden coupling, enforce predictable behavior, and make logic testable by controlling scope and organizing code into well-designed procedures and classes.

## Scope Rules

### Prefer Small Local Scope

```vb
Public Sub ProcessData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
End Sub
```
Avoid module-level variables unless necessary.

### When to Use Module-Level Private

- Shared helper state for 2-3 functions.
- Cached objects that are expensive to access.

### Never Use Public Variables

Global state creates fragile, unpredictable code.

## Encapsulation with Classes

### Example: Class `clsCustomer`

```vb
Private mId As String
Private mBalance As Currency

Public Property Get Id() As String: Id = mId: End Property
Public Property Let Id(ByVal value As String)
    If Len(value) = 0 Then Err.Raise 1001, , "Id cannot be empty"
    mId = value
End Property
```

## Clear Procedure Design

### Guard Clauses

```vb
If ws Is Nothing Then Exit Sub
If Not IsValidData(arr) Then Exit Sub
```

### Long Procedures → Pipeline

```vb
Public Sub RunWorkflow()
    LoadData
    ValidateData
    TransformData
    WriteOutput
End Sub
```

## Do / Don't

### Do

- Break complex logic into multiple functions.
- Document expectations and invariants.
- Use classes for domain concepts.

### Don't

- Allow 500-line procedures.
- Mutate parameters accidentally via `ByRef`.

## Anti-Patterns

BAD: Hidden shared state.
BAD: Classes with only public fields.
BAD: Procedures dependent on module-level scratch variables.

## Acceptance Checklist

- [ ] All public state justified.
- [ ] No procedure longer than one screen.
- [ ] Domain logic encapsulated.
