Attribute VB_Name = "ModMain"
'===============================================================
' Module ModMain
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Public Function NoRows() As Integer
    NoRows = Application.WorksheetFunction.CountA(ShtMain.Range(RNG_CREW_COUNT)) - 1
End Function

Public Sub PerfSettingsOn()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Public Sub PerfSettingsOff()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub


