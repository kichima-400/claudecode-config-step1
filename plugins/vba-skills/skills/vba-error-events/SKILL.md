---
name: vba-error-events-team
description: Robust error handling, event safety, debugging, logging.
---
# Error Handling, Events & Team Practices — Expanded Edition

*Ensure reliability, traceability, and predictable execution across users and environments.*

## Purpose

Create VBA applications that fail safely, provide diagnostic information, prevent re-entrant chaos, and support professional deployment workflows.

## Structured Error Handling

### Standard Template

```vb
Public Sub RunTask()
    On Error GoTo CleanFail
    SaveState
    ' Main work
CleanExit:
    RestoreState
    Exit Sub
CleanFail:
    LogError "RunTask", Err
    Resume CleanExit
End Sub
```

### Use Narrow Resume Next

```vb
On Error Resume Next
Set ws = wb.Worksheets(name)
If Err.Number <> 0 Then
    Err.Clear
    ' Handle missing sheet
End If
On Error GoTo 0
```

## Debugging & Diagnostics

### Use Immediate Window

```vb
Debug.Print "Processing row", i
```

### Assertions

```vb
Debug.Assert Not ws Is Nothing
```

### Centralized Logging

```vb
Public Sub LogError(ByVal proc As String, ByVal er As ErrObject)
    Debug.Print proc, er.Number, er.Description
End Sub
```

## Event Safety

### Prevent Re-Entrancy

```vb
If mHandlingEvent Then Exit Sub
mHandlingEvent = True
' ...
mHandlingEvent = False
```

### Keep Events Thin

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanFail
    If Not Intersect(Target, Me.Range("Data")) Is Nothing Then
        HandleDataChange Target
    End If
CleanExit:
    Exit Sub
CleanFail:
    LogError "Worksheet_Change", Err
    Resume CleanExit
End Sub
```

## Anti-Patterns

BAD: Silent failure via blanket `Resume Next`.
BAD: Unlogged runtime errors.
BAD: Event procedures containing business logic.

## Acceptance Checklist

- [ ] Error handler in every non-trivial procedure.
- [ ] Event code non-reentrant.
- [ ] Logging and diagnostics in place.
- [ ] Deployment steps documented.
