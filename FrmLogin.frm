VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmLogin 
   Caption         =   "Enter Date"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "FrmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'===============================================================
' Module FrmLogin
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Private SelUser As EnumUserLvl

Public Function ShowForm() As EnumUserLvl
    SelUser = DevLvl
    Show
    ShowForm = SelUser
End Function

Private Sub BtnOK_Click()
    Unload Me
End Sub

Private Sub OptAdmin_Click()
    SelUser = AdminLvl
End Sub

Private Sub OptDev_Click()
    SelUser = DevLvl
End Sub

Private Sub OptNormal_Click()
    SelUser = BasicLvl
End Sub

Private Sub UserForm_Initialize()
    OptDev.Value = True
End Sub
