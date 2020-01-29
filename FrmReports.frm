VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReports 
   Caption         =   "Run Reports"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "FrmReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelReport As EnumReport

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnRunExpQual_Click()
    ModReports.ExpQualReport
End Sub

Private Sub BtnRunProm_Click()
    If SelReport = 0 Then
        MsgBox "Please select a Report", vbInformation + vbOKOnly
    Else
        ModReports.PromReports SelReport
    End If
End Sub

Private Sub OptPromAC_Click()
    SelReport = SCtoAC
End Sub

Private Sub OptPromCM_Click()
    SelReport = DOtoCM
End Sub

Private Sub OptPromDO_Click()
    SelReport = FFtoDO
End Sub

Private Sub OptPromSC_Click()
    SelReport = CMtoSC
End Sub

