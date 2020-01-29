VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmImpExp 
   Caption         =   "Data Import and Export"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2850
   OleObjectBlob   =   "FrmImpExp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmImpExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnExport_Click()
    ModDataImpExp.ExportData
End Sub

