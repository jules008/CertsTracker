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
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Private Sub Workbook_Open()
    InitialiseSystem
End Sub

Public Sub InitialiseSystem()
    PerfSettingsOn
    DetectUser
    SelectUser
    If USER_LEVEL = DevLvl Then DevView
    If USER_LEVEL = AdminLvl Then AdminView
    If USER_LEVEL = BasicLvl Then BasicView
    ShtMain.RefreshQuals
    PerfSettingsOff
End Sub

Public Sub SelectUser()
    If USER_LEVEL = DevLvl Then USER_LEVEL = FrmLogin.ShowForm
End Sub
