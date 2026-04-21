# 📊 VBA Excel Cleaner

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![VBA](https://img.shields.io/badge/VBA-Excel-orange.svg)
![Excel](https://img.shields.io/badge/Microsoft%20Excel-Enabled-brightgreen.svg)

------------------------------------------------------------------------

## 🎯 Objectif

Automatiser le nettoyage d'un classeur Excel : - supprimer toutes les
formules - supprimer les objets/formulaires - figer les données pour
export ou archivage

------------------------------------------------------------------------

## ⚙️ Suppression des formules

``` vb
Sub SupprimerFormulesToutesFeuilles()

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

    MsgBox "Formules supprimées", vbInformation

End Sub
```

------------------------------------------------------------------------

## 🧹 Suppression des objets

``` vb
Sub SupprimerFormulairesToutesFeuilles()

    Dim ws As Worksheet
    Dim shp As Shape

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Objets supprimés", vbInformation

End Sub
```

------------------------------------------------------------------------

## 🚀 Installation

1.  ALT + F11
2.  Insertion \> Module
3.  Coller le code
4.  Sauvegarder en .xlsm

------------------------------------------------------------------------

## ▶️ Utilisation

ALT + F8 → choisir la macro → Exécuter

------------------------------------------------------------------------

## ⚠️ Précaution

Cette action est irréversible. Faites une copie du fichier.

------------------------------------------------------------------------
