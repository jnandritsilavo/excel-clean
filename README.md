# VBA Excel Cleaner

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![VBA](https://img.shields.io/badge/VBA-Excel-orange.svg)
![Excel](https://img.shields.io/badge/Microsoft%20Excel-Enabled-brightgreen.svg)

A small set of VBA macros that strips a workbook down to its raw values:
formulas are converted to static data and every embedded object is removed.

---

## Purpose

An Excel workbook that is still "alive" carries formulas, external links,
controls and embedded objects. That is fine while the file is being worked on,
but it becomes a liability once the file leaves its original context.

Use this tool when you need to:

- **Archive a workbook.** Frozen values stay readable years later, even if the
  source files, add-ins or named ranges they depended on are gone.
- **Share a file outside the team.** Formulas can expose calculation logic,
  pricing rules or internal references. Removing them ships the results only.
- **Break external links.** A workbook that points to other files shows
  update prompts and `#REF!` errors on the recipient's machine.
- **Publish a stable snapshot.** Monthly reports, invoices and figures sent for
  validation must not change when the workbook is reopened or recalculated.
- **Reduce size and improve performance.** Removing volatile formulas and
  leftover shapes makes large workbooks noticeably lighter and faster to open.
- **Clean up an inherited file.** Old buttons, orphaned form controls and
  invisible objects accumulate over time and clutter the sheets.

If you still need the workbook to compute, this tool is not what you want:
the operation is one-way.

---

## Requirements

- Microsoft Excel with VBA enabled (Windows or macOS)
- A workbook saved in a macro-enabled format (`.xlsm`)
- Macro execution allowed in the Trust Center

---

## Removing formulas — `RemoveFormulasAllSheets`

Loops through every worksheet and replaces the used range with its own values.
Screen updating, automatic calculation and events are switched off during the
operation, then restored.

```vb
Sub RemoveFormulasAllSheets()

    Dim ws As Worksheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For Each ws In ThisWorkbook.Worksheets
        If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
            ws.UsedRange.Value = ws.UsedRange.Value
        End If
    Next ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Formulas removed", vbInformation

End Sub
```

The `CountA` check skips empty sheets, which avoids unnecessary work and
keeps the macro fast on workbooks with many unused tabs.

---

## Removing objects — `RemoveShapesAllSheets`

Deletes every shape on every worksheet: buttons, form controls, ActiveX
controls, text boxes, images, charts and grouped objects.

```vb
Sub RemoveShapesAllSheets()

    Dim ws As Worksheet
    Dim shp As Shape

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Objects removed", vbInformation

End Sub
```

Note that this removes charts and pictures as well, not only form controls.
Run it only when a plain data workbook is the intended result.

---

## Installation

1. Press `ALT + F11` to open the VBA editor.
2. Choose `Insert > Module`.
3. Paste the code into the module.
4. Save the workbook as `.xlsm`.

---

## Usage

1. Press `ALT + F8`.
2. Select the macro to run.
3. Click `Run`.

Run `RemoveFormulasAllSheets` first, then `RemoveShapesAllSheets`, so that any
control still bound to a cell is removed after the values have been frozen.

---

## Warning

Both operations are irreversible and cannot be undone with `CTRL + Z`.
Work on a copy of the file, or save the original under a different name before
running the macros.

---

## License

Released under the MIT License.
