VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtMain
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 04 Feb 20
'===============================================================
Option Explicit

Public Sub RefreshQuals(Optional RowNo As Integer)
    Dim RngCell As Range
    Dim RngRoleQualArea As Range
    Dim RequiredYN As EnumTriState
    Dim GainedYN As EnumTriState
    Dim Role As EnumRole
    Dim Qual As EnumQual
    Dim Status As EnumStatus
    Dim QualStatus As EnumExpiryStatus
    Dim SSN As String
    Dim CourseDate As String
    Dim NoCrew As Integer
    
    On Error GoTo ErrorHandler
    
    NoCrew = ModMain.NoRows
    
    If RowNo <> 0 Then
        Set RngRoleQualArea = Me.Range(RNG_WORKING_START & RowNo & RNG_WORKING_END & RowNo)
    Else
        Set RngRoleQualArea = Me.Range(RNG_WORKING & NoCrew + 3)
    End If
    
    ModMain.PerfSettingsOn
    
    For Each RngCell In RngRoleQualArea
        
        Role = GetRole(RngCell.Row)
        Qual = GetQual(RngCell.Column)
        SSN = GetSSN(RngCell.Row)
        Status = GetStatus(RngCell.Row)
        CourseDate = ShtCourseDates.LookUpCourseDate(SSN, Qual, eRead)
        RequiredYN = ShtRoleLU.LookUpQual(Role, Qual)
        QualStatus = ShtRoleLU.RetQualStatus(CourseDate, Qual)
        
        If CourseDate = "" Then
            GainedYN = No
        Else
            GainedYN = Yes
        End If
        
        Select Case RequiredYN
            Case Is = Yes
                If GainedYN = Yes Then
                    If Status = Active Then
                        If QualStatus = Valid Then RngCell = 4
                        If QualStatus = Expired Then RngCell = -5
                    End If
                    If Status = inactive Then RngCell = 9
                Else
                    If Status = Active Then RngCell = 3
                    If Status = inactive Then RngCell = 8
                End If
            Case Is = No
                If GainedYN = Yes Then
                    If Status = Active Then
                        If QualStatus = Valid Then RngCell = 1
                        If QualStatus = Expired Then RngCell = -2
                    End If
                    If Status = inactive Then RngCell = 7
                Else
                    If Status = Active Then RngCell = 0
                    If Status = inactive Then RngCell = 6
                End If
            Case Is = Blank
                RngCell = ""
        End Select
        
'        Debug.Print RngCell.Address, "Required - " & RequiredYN, "Gained - " & GainedYN, "status - " & Status, "Value - " & RngCell.Value
    Next
        
    ModMain.PerfSettingsOff
Exit Sub

ErrorHandler:
    
    ModMain.PerfSettingsOff

End Sub

Private Function GetSSN(RowNo As Integer) As String
    GetSSN = Me.Cells(RowNo, 6)
End Function

Private Function GetStatus(RowNo As Integer) As EnumStatus
    If Me.Cells(RowNo, 7) = "Active" Then GetStatus = Active
    If Me.Cells(RowNo, 7) = "Inactive" Then GetStatus = inactive
End Function


Private Function GetRole(RowNo As Integer) As EnumRole
    Dim StrRole As String
    
    StrRole = Me.Cells(RowNo, 3)
    
    If StrRole <> "" Then
        Select Case StrRole
            Case "FI"
                GetRole = FI
            Case "Dispatch"
                GetRole = Dispatch
            Case "Firefighter"
                GetRole = Firefighter
            Case "Driver/Op"
                GetRole = DriverOp
            Case "Crew Manager"
                GetRole = CrewManager
            Case "Station Captain"
                GetRole = StationCaptain
            Case "A/C Training"
                GetRole = ACTraining
            Case "A/C Health and Safety"
                GetRole = ACHealthandSafety
            Case "A/C Fire Prevention"
                GetRole = ACFirePrevention
            Case "A/C Ops"
                GetRole = ACOps
            Case "Deputy Chief"
                GetRole = DeputyChief
            Case "Fire Chief"
                GetRole = FireChief
            Case Else
                MsgBox "cannot find role", vbCritical
        End Select
    End If
End Function

Private Function GetQual(ColNo As Integer) As EnumQual
    Dim StrQual As String
    
    GetQual = Me.Cells(1, ColNo)

End Function

Private Sub BtnAddNew_Click()
    PerfSettingsOn
    Unprotect "2683174"
    Range("G" & Me.GetCrewCountAll + 4) = "Active"
    Range("G" & Me.GetCrewCountAll + 4).Select
    ShowHide
    ShowHide
    Protect "2683174"
    PerfSettingsOff
End Sub

Private Sub BtnEnterDate_Click()
    Dim SSN As String
    
    If Intersect(ActiveCell, Range("A4:A1000")) Is Nothing Then
        SSN = GetSSN(ActiveCell.Row)
        
        If SSN = "" Then Exit Sub
        
        FrmTngDates.ShowForm SSN
    End If
End Sub

Private Sub BtnImpExp_Click()
    FrmImpExp.Show
End Sub

Private Sub CmdRefresh_Click()
    ThisWorkbook.InitialiseSystem
End Sub

Private Sub CmdReports_Click()
    FrmReports.Show
End Sub

Private Sub CmdShowHide_Click()
    ShowHide
End Sub

Private Sub CmdSort_Click()
    SortBy CmoSortby.ListIndex
End Sub

Private Sub Worksheet_Activate()
    PerfSettingsOff
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim CrewCnt As Integer
    Dim SSN As String
    Dim Qual As EnumQual
       
    CrewCnt = GetCrewCountAll
    SSN = GetSSN(Target.Row)
    Qual = GetQual(Target.Column)
    
    If USER_LEVEL = BasicLvl Then Exit Sub
    
    If Not Intersect(Target, Me.Range(RNG_WORKING & CrewCnt + 3)) Is Nothing Then
        
        Debug.Print "FrmEnterDate.ShowForm", SSN, Qual
        FrmEnterDate.ShowForm SSN, Qual

        Me.RefreshQuals Target.Row
    End If
    Target.Offset(1, 0).Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim NoRows As Integer
    
    NoRows = ModMain.NoRows
    
    If Target.Count > 1 Then Exit Sub
    
    If Not Intersect(Target, Me.Range(RNG_PERSON_DET_START & "3:" & RNG_PERSON_DET_END & NoRows + 3)) Is Nothing Then
        Me.RefreshQuals Target.Row
    End If
    
    If Not Intersect(Target, Me.Range(RNG_ACTIVE)) Is Nothing Then
        Me.RefreshQuals Target.Row
    End If
    
End Sub


Public Function GetCrewCountAll() As Integer
    GetCrewCountAll = Application.WorksheetFunction.CountA(ShtMain.Range("A:A")) - 1
End Function

Public Function GetName(SSN As String) As String
    Dim RngSSN As Range
    Dim RngResult As Range
    
    Set RngSSN = Me.Range("F:F")
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then Exit Function
    
    GetName = RngResult.Offset(0, -5)
    
    Set RngResult = Nothing
    Set RngSSN = Nothing
    
End Function

Public Sub SortBy(Field As Integer)
    Dim RngSort As Range
    Dim RngIndex As Range
    
    On Error GoTo ErrorHandler
    
    Set RngSort = Range("A4:AS" & ModMain.NoRows + 3)
    Set RngIndex = Range("A4")
    
    Me.Unprotect "2683174"

    RngSort.Sort RngIndex.Offset(0, Field)
    
    If USER_LEVEL <> DevLvl Then Me.Protect "2683174"
    
    Set RngSort = Nothing
    Set RngIndex = Nothing
    
    Exit Sub
    
ErrorHandler:

    If USER_LEVEL <> DevLvl Then Me.Protect "2683174"
    
    Set RngSort = Nothing
    Set RngIndex = Nothing

End Sub

Private Sub ShowHide()
    Dim RngShowHide As Range
    
    On Error GoTo ErrorHandler
    
    Set RngShowHide = Range("A3:AS" & ModMain.NoRows + 3)
    
    Me.Unprotect "2683174"
    
    If CmdShowHide.Caption = "Hide Leavers" Then
        RngShowHide.AutoFilter Field:=7, Criteria1:="Active", VisibleDropDown:=False
        CmdShowHide.Caption = "Show Leavers"
    Else
        ShtMain.AutoFilterMode = False
        CmdShowHide.Caption = "Hide Leavers"
    End If
    
    If USER_LEVEL <> DevLvl Then Me.Protect "2683174"
    
    Set RngShowHide = Nothing
    
    Exit Sub
    
ErrorHandler:

    If USER_LEVEL <> DevLvl Then Me.Protect "2683174"
    
    Set RngShowHide = Nothing
End Sub

Public Function GetDataAll() As Variant()
    Dim AryDataAll() As Variant
    AryDataAll = Range("A4:AS" & GetCrewCountAll + 3)
    GetDataAll = AryDataAll
End Function

Public Function GetPersDetails() As Variant()
    Dim AryPersDet As Variant
    AryPersDet = Range("A4:G" & GetCrewCountAll + 3)
    GetPersDetails = AryPersDet
End Function

Public Sub ClearPersDetails()
    Range("A4:G" & GetCrewCountAll + 3).ClearContents
End Sub


