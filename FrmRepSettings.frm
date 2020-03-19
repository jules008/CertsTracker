VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmRepSettings 
   Caption         =   "Report Settings"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   OleObjectBlob   =   "FrmRepSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmRepSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmRepSettings
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Mar 20
'===============================================================
Option Explicit

Private Watch As String
Private Period As Integer
Private PeriodSel As Boolean

'===============================================================
' BtnRun_Click
' checks button selections and runs report
'---------------------------------------------------------------
Private Sub BtnRun_Click()
    If Validation Then ModReports.ExpQualReport Watch, Period
    Unload Me
End Sub

'===============================================================
' OptAll_Click
'---------------------------------------------------------------
Private Sub OptAll_Click()
    Watch = "All"
End Sub

'===============================================================
' OptBlue_Click
'---------------------------------------------------------------
Private Sub OptBlue_Click()
    Watch = "Blue"
End Sub

'===============================================================
' OptExp10_Click
'---------------------------------------------------------------
Private Sub OptExp10_Click()
    Period = 10
    PeriodSel = True
End Sub

'===============================================================
' OptExp30_Click
'---------------------------------------------------------------
Private Sub OptExp30_Click()
    Period = 30
    PeriodSel = True
End Sub

'===============================================================
' OptExp60_Click
'---------------------------------------------------------------
Private Sub OptExp60_Click()
    Period = 60
    PeriodSel = True
End Sub

'===============================================================
' OptRed_Click
'---------------------------------------------------------------
Private Sub OptRed_Click()
    Watch = "Red"
End Sub

'===============================================================
' OptWhite_Click
'---------------------------------------------------------------
Private Sub OptWhite_Click()
    Watch = "White"
End Sub

'===============================================================
' OptExp0_Click
'---------------------------------------------------------------
Private Sub OptExp0_Click()
    Period = 0
    PeriodSel = True
End Sub

'===============================================================
' Validation
' Ensures both sets of radio buttons are selecting before proceeding
'---------------------------------------------------------------
Private Function Validation() As Boolean
    If Watch = "" Or PeriodSel = False Then Validation = False Else Validation = True
End Function

