VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtRoleLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtRoleLU
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - Added Days to expiry function
' v1.1.1 - Added global ranges
'---------------------------------------------------------------
' Date - 23 Mar 20
'===============================================================
Option Explicit

Private Const RNG_QUALS = "B2:AN2"
'===============================================================
' LookUpQual
' Checks to see whether a particular qualification is required for a role
'---------------------------------------------------------------
Public Function LookUpQual(Role As EnumRole, Qual As EnumQual) As EnumTriState
    
    If Role = 0 Then
        LookUpQual = Blank
    Else
        If Me.Cells(Role + 2, Qual + 1).Value = 1 Then
            LookUpQual = Yes
        Else
            LookUpQual = No
        End If
    End If
End Function

'===============================================================
' RetQualStatus
' Returns the status of a qualification, either valid or expired
'---------------------------------------------------------------
Public Function RetQualStatus(QualDate As Variant, Qual As EnumQual) As EnumExpiryStatus ' detect if exam has expired
    Dim QualDur As Integer
    Dim LimitedQual As Boolean
    Dim LUValue As Integer
    
    LUValue = Me.Cells(21, Qual + 1).Value
    
    If IsNumeric(QualDate) And QualDate > 10000 Then QualDate = CDate(QualDate)
    
    If IsDate(QualDate) Then
        If LUValue <> 0 Then LimitedQual = True Else LimitedQual = False
        If LimitedQual Then
            QualDur = LUValue
            
            If DateAdd("m", QualDur, QualDate) < Now Then
                RetQualStatus = Expired
            Else
                RetQualStatus = Valid
            End If
        Else
            RetQualStatus = Valid
        End If
    Else
        If QualDate = "1" Then RetQualStatus = Valid
        If QualDate = "2" Then RetQualStatus = Valid
        If QualDate = "EMT" Then RetQualStatus = Valid
    End If
    
End Function

'===============================================================
' DaysTillExpiry
' Takes a date and returns how many days till expiry.  If not
' applicable, returns 999
'---------------------------------------------------------------
Public Function DaysTillExpiry(QualDate As Variant, Qual As EnumQual) As Integer
    Dim QualDur As Integer
    Dim QDate As Date
    
    QualDur = Me.Cells(21, Qual + 1).Value
    
    If QualDur = 0 Or _
      QualDate = "EMT" Or _
      QualDate = 1 Then
        DaysTillExpiry = 999
    Else
        QDate = CDate(QualDate)
        QDate = DateAdd("m", QualDur, QDate)
        DaysTillExpiry = QDate - Now + 1
        
    End If
End Function

'===============================================================
' GetRoleEligibility
' Returns an array with the quals required for a specified role
'---------------------------------------------------------------
Public Function GetRoleEligibility(Role As EnumRole) As Variant()
    GetRoleEligibility = Range(RNG_QUALS).Offset(Role, 0)
End Function








