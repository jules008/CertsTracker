VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmShowHideCols 
   Caption         =   "Show / Hide Columns"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2670
   OleObjectBlob   =   "FrmShowHideCols.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmShowHideCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmShowHideCols
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 21 Mar 20
'===============================================================
Option Explicit

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    ShowHideCols
    Unload Me
End Sub

'===============================================================
' UserForm_Initialize
'---------------------------------------------------------------
Private Sub UserForm_Initialize()
    SetBoxes
End Sub

'===============================================================
' ShowHideCols
' Shows and hides columns as selected on event
'---------------------------------------------------------------
Private Sub ShowHideCols()
    With ShtMain
        .Unprotect SEC_KEY
        
        If ChkDoB Then .SetColStatus eDoB, True Else .SetColStatus eDoB, False
        If ChkFIN Then .SetColStatus eFINNo, True Else .SetColStatus eFINNo, False
        If ChhDoDRef Then .SetColStatus eDoDRef, True Else .SetColStatus eDoDRef, False
        
        .SetControlPos "CmoSortBy", 102, 252, 19, 80
        .SetControlPos "CmdSort", 102, 330, 18, 33
        
        If USER_LEVEL <> DevLvl Then .Protect SEC_KEY
    End With
End Sub

'===============================================================
' SetBoxes
' Reads the status of the columns and sets the check boxes
' accordingly
'---------------------------------------------------------------
Private Sub SetBoxes()
    With ShtMain
        ChkDoB = .ReadColStatus(eDoB)
        ChkFIN = .ReadColStatus(eFINNo)
        ChhDoDRef = .ReadColStatus(eDoDRef)
    End With
End Sub
