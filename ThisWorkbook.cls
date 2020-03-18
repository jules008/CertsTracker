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
'---------------------------------------------------------------
' Date - 05 Mar 20
'===============================================================
Option Explicit

Private Sub Workbook_Open()
    InitialiseSystem
End Sub


