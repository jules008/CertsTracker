VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ThisWorkbook
'===============================================================
' v1.0.0 - Initial Version
' v1.0.1 - Moved start up functions to ModStartup
' v1.1.0 - Before and after save events
'---------------------------------------------------------------
' Date - 19 Mar 20
'===============================================================
Option Explicit

'===============================================================
' Workbook_AfterSave
'---------------------------------------------------------------
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    InitialiseSystem False
End Sub

'===============================================================
' Workbook_BeforeSave
'---------------------------------------------------------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    USER_LEVEL = BasicLvl
    ModSecurity.BasicView
End Sub

'===============================================================
' Workbook_Open
'---------------------------------------------------------------
Private Sub Workbook_Open()
    InitialiseSystem False
End Sub


