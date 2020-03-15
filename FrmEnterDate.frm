VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEnterDate 
   Caption         =   "Enter Date"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   OleObjectBlob   =   "FrmEnterDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmEnterDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmEnterDate
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Private SSN As String
Private Qual As EnumQual

Public Sub ShowForm(LocSSN As String, LocQual As EnumQual)
    SSN = LocSSN
    Qual = LocQual
    
    LblName = "Name: " & ShtMain.GetName(SSN)
    LblQual = "Qual: " & QualConvEnum(Qual)

    TxtDate = Format(Now, "dd mmm yy")
    Show
End Sub

Private Sub BtnClearDate_Click()
    If Selection.Cells.Count > 1 Then Exit Sub
    
    ShtCourseDates.LookUpCourseDate SSN, Qual, EClear
    Unload Me
End Sub

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub BtnOK_Click()
    Dim CourseDate As String
    
    If TxtDate = "" Then Exit Sub
    
    If Not IsDate(TxtDate) Then Exit Sub
        
    If Qual = EMR Then
        CourseDate = ShtCourseDates.LookUpCourseDate(SSN, EMT, eRead)
        If CourseDate <> "" Then TxtDate = "EMT"
    End If
    
    ShtCourseDates.LookUpCourseDate SSN, Qual, EWrite, TxtDate
    
    Unload Me
End Sub
