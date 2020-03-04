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
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

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

Public Function RetQualStatus(QualDate As Variant, Qual As EnumQual) As EnumExpiryStatus ' detect if exam has expired
    Dim QualDur As Integer
    Dim LimitedQual As Boolean
    Dim LUValue As Integer
    
    LUValue = Me.Cells(21, Qual + 1).Value
    
    Debug.Print "Qual Date - " & QualDate, RetQualStatus
    
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
