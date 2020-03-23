VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTngDates 
   Caption         =   "Training Dates"
   ClientHeight    =   9300
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
' v1.1.0 - Added reference matrix for control look up
'---------------------------------------------------------------
' Date - 23 Mar 20
'===============================================================
Option Explicit

Dim SSN As String

'===============================================================
' PopulateForm
'---------------------------------------------------------------
Private Sub PopulateForm()
    Dim AryDates() As Variant
    Dim Name As String
    Dim i As Integer
    Dim Course As EnumQual
    Dim CntrlName As String
    On Error GoTo ErrorHandler
    
    AryDates = ShtCourseDates.GetAllDates(SSN)
    Name = ShtMain.GetName(SSN)
    LblName = Name
    
    For i = 1 To NO_COURSES
        Course = i
        CntrlName = GetControlName(Course)
        
        If AryDates(1, i) = 1 Then
            Me.Controls(CntrlName).Value = "Passed"
        Else
            Me.Controls(CntrlName).Value = Format(AryDates(1, i), "dd mmm yy")
        End If
    Next
Exit Sub

ErrorHandler:
    Debug.Print "Unable to retrieve course dates"
End Sub

'===============================================================
' ShowForm
'---------------------------------------------------------------
Public Sub ShowForm(LocSSN As String)
    SSN = LocSSN
    
    PopulateForm
    Show
End Sub

'===============================================================
' BtnClose_Click
'---------------------------------------------------------------
Private Sub BtnClose_Click()
    Unload Me
End Sub

'===============================================================
' GetControlName
' matrix to return the control name from given course
'---------------------------------------------------------------
Private Function GetControlName(ByRef Course As EnumQual) As String
    Select Case Course
        Case CPR
            GetControlName = "TxtDates1"
        Case PPEProgram
            GetControlName = "TxtDates38"
        Case EMR
            GetControlName = "TxtDates2"
        Case Munitions
            GetControlName = "TxtDates3"
        Case IS100_IS700
            GetControlName = "TxtDates4"
        Case IS200_IS800
            GetControlName = "TxtDates5"
        Case HazmatAW
            GetControlName = "TxtDates6"
        Case HazmatOps
            GetControlName = "TxtDates7"
        Case FirefighterI
            GetControlName = "TxtDates8"
        Case FirefighterII
            GetControlName = "TxtDates9"
        Case TelecommunicatorI
            GetControlName = "TxtDates10"
        Case TelecommunicatorII
            GetControlName = "TxtDates11"
        Case LGVCatC
            GetControlName = "TxtDates12"
        Case DrvrOpPumper
            GetControlName = "TxtDates13"
        Case DrvrOPMWS
            GetControlName = "TxtDates14"
        Case HazmatTech
            GetControlName = "TxtDates15"
        Case FireOfficerI
            GetControlName = "TxtDates16"
        Case FireInpsectorI
            GetControlName = "TxtDates17"
        Case FireInstructorI
            GetControlName = "TxtDates18"
        Case IncidentSafetyOfficer
            GetControlName = "TxtDates19"
        Case FireOfficerII
            GetControlName = "TxtDates20"
        Case FireInspectorII
            GetControlName = "TxtDates21"
        Case FireInstructorII
            GetControlName = "TxtDates22"
        Case HazmatIC
            GetControlName = "TxtDates23"
        Case NIMS300400
            GetControlName = "TxtDates24"
        Case FireOfficerIII
            GetControlName = "TxtDates25"
        Case FireInspectorIII
            GetControlName = "TxtDates26"
        Case FireInstructorIII
            GetControlName = "TxtDates27"
        Case FireOfficerIV
            GetControlName = "TxtDates28"
        Case EMT
            GetControlName = "TxtDates29"
        Case HealthSafetyOfficer
            GetControlName = "TxtDates30"
        Case HazmatWMDIC
            GetControlName = "TxtDates31"
        Case RescueTechnicianI
            GetControlName = "TxtDates32"
        Case RescueTechnicianII
            GetControlName = "TxtDates33"
        Case PlansExaminer
            GetControlName = "TxtDates34"
        Case MSASCBAServicer
            GetControlName = "TxtDates35"
        Case WMD
            GetControlName = "TxtDates36"
        Case LGVCatCE
            GetControlName = "TxtDates37"
        Case FireEducator
            GetControlName = "TxtDates39"
    End Select
End Function
