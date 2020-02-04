VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTngDates 
   Caption         =   "Training Dates"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   OleObjectBlob   =   "FrmTngDates.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTngDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmTrngDates
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Dim SSN As String

Private Sub PopulateForm()
    Dim AryDates() As Variant
    Dim Name As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    AryDates = ShtCourseDates.GetAllDates(SSN)
    Name = ShtMain.GetName(SSN)
    
    LblName = Name
    
    For i = 1 To NO_COURSES
        If AryDates(1, i) = 1 Then
            Me.Controls("TxtDates" & i).Value = "Passed"
        Else
            Me.Controls("TxtDates" & i).Value = Format(AryDates(1, i), "dd mmm yy")
        End If
    Next
Exit Sub

ErrorHandler:
    Debug.Print "Unable to retrieve course dates"
End Sub

Public Sub ShowForm(LocSSN As String)
    SSN = LocSSN
    
    PopulateForm
    Show
End Sub

Private Sub BtnClose_Click()
    Unload Me
End Sub

