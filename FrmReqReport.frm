VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReqReport 
   Caption         =   "Required Qualification"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   OleObjectBlob   =   "FrmReqReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReqReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' Module FrmReqReport
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Apr 20
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
    If Validation Then
        ModReports.ReqQualReport Watch, CmoQuals.ListIndex + 1
        Unload Me
    Else
        MsgBox "Please check the form selections", vbOKOnly & vbExclamation, APP_NAME
    End If
    
    
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
' Validation
' Ensures both sets of radio buttons are selecting before proceeding
'---------------------------------------------------------------
Private Function Validation() As Boolean
    If Watch = "" Or CmoQuals.ListIndex = -1 Then Validation = False Else Validation = True
End Function

'===============================================================
' UserForm_Initialize
'---------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim Entry As Range
    
    With CmoQuals
        For Each Entry In ShtLists.Range("LU_COURSES")
            .AddItem Entry
        Next
    
    End With
End Sub
