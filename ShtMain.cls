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
' v1.0.1 - formatting changes
'---------------------------------------------------------------
' Date - 17 Mar 20
'===============================================================
Option Explicit

Private Const RNG_PERS_DET As String = "A4:G"
' ===============================================================
' RefreshQuals
' Refreshes all qulaification data.  If Row No is specified, only
' refreshes that line.
' ---------------------------------------------------------------
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

' ===============================================================
' GetSSN
' Returns SSN for active line
' ---------------------------------------------------------------
Private Function GetSSN(RowNo As Integer) As String
    GetSSN = Me.Cells(RowNo, 6)
End Function

' ===============================================================
' GetStatus
' Returns Status for active line
' ---------------------------------------------------------------
Private Function GetStatus(RowNo As Integer) As EnumStatus
    If Me.Cells(RowNo, 7) = "Active" Then GetStatus = Active
    If Me.Cells(RowNo, 7) = "Inactive" Then GetStatus = inactive
End Function

' ===============================================================
' GetRole
' Returns Role for active line
' ---------------------------------------------------------------
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

' ===============================================================
' GetQual
' Returns Qualification for active Col
' ---------------------------------------------------------------
Private Function GetQual(ColNo As Integer) As EnumQual
    Dim StrQual As String
    
    GetQual = Me.Cells(1, ColNo)

End Function

' ===============================================================
' BtnAddNew_Click
' ---------------------------------------------------------------
Private Sub BtnAddNew_Click()
    PerfSettingsOn
    Unprotect SEC_KEY
    Range("G" & Me.GetCrewCountAll + 4) = "Active"
    Range("G" & Me.GetCrewCountAll + 4).Select
    ShowHide
    ShowHide
    Protect SEC_KEY
    PerfSettingsOff
End Sub

' ===============================================================
' BtnEnterDate_Click
' ---------------------------------------------------------------
Private Sub BtnEnterDate_Click()
    Dim SSN As String
    
    If Intersect(ActiveCell, Range("A4:A1000")) Is Nothing Then
        SSN = GetSSN(ActiveCell.Row)
        
        If SSN = "" Then Exit Sub
        
        FrmTngDates.ShowForm SSN
    End If
End Sub

' ===============================================================
' BtnImpExp_Click
' ---------------------------------------------------------------
Private Sub BtnImpExp_Click()
    FrmImpExp.Show
End Sub

' ===============================================================
' CmdRefresh_Click
' ---------------------------------------------------------------
Private Sub CmdRefresh_Click()
    ThisWorkbook.InitialiseSystem
End Sub

' ===============================================================
' CmdReports_Click
' ---------------------------------------------------------------
Private Sub CmdReports_Click()
    FrmReports.Show
End Sub

' ===============================================================
' CmdShowHide_Click
' ---------------------------------------------------------------
Private Sub CmdShowHide_Click()
    ShowHide
End Sub

' ===============================================================
' CmdSort_Click
' ---------------------------------------------------------------
Private Sub CmdSort_Click()
    SortBy CmoSortby.ListIndex
End Sub

' ===============================================================
' Worksheet_Activate
' ---------------------------------------------------------------
Private Sub Worksheet_Activate()
    PerfSettingsOff
End Sub

' ===============================================================
' Worksheet_BeforeDoubleClick
' ---------------------------------------------------------------
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

' ===============================================================
' Worksheet_Change
' ---------------------------------------------------------------
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

' ===============================================================
' GetCrewCountAll
' Returns total number of people
' ---------------------------------------------------------------

Public Function GetCrewCountAll() As Integer
    GetCrewCountAll = Application.WorksheetFunction.CountA(ShtMain.Range("A:A")) - 1
End Function

' ===============================================================
' GetName
' Returns name from SSN
' ---------------------------------------------------------------
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

' ===============================================================
' SortBy
' sorts personal details by submitted field
' ---------------------------------------------------------------
Public Sub SortBy(Field As Integer)
    Dim RngSort As Range
    Dim RngIndex As Range
    
    On Error GoTo ErrorHandler
    
    Set RngSort = Range("A4:AS" & ModMain.NoRows + 3)
    Set RngIndex = Range("A4")
    
    Me.Unprotect SEC_KEY

    RngSort.Sort RngIndex.Offset(0, Field)
    
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    
    Set RngSort = Nothing
    Set RngIndex = Nothing
    
    Exit Sub
    
ErrorHandler:

    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    
    Set RngSort = Nothing
    Set RngIndex = Nothing

End Sub

' ===============================================================
' ShowHide
' alternately shows or hides active people.
' ---------------------------------------------------------------
Private Sub ShowHide()
    Dim RngShowHide As Range
    
    On Error GoTo ErrorHandler
    
    Set RngShowHide = Range("A3:AS" & ModMain.NoRows + 3)
    
    Me.Unprotect SEC_KEY
    
    If CmdShowHide.Caption = "Hide Leavers" Then
        RngShowHide.AutoFilter Field:=7, Criteria1:="Active", VisibleDropDown:=False
        CmdShowHide.Caption = "Show Leavers"
    Else
        ShtMain.AutoFilterMode = False
        CmdShowHide.Caption = "Hide Leavers"
    End If
    
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    
    Set RngShowHide = Nothing
    
    Exit Sub
    
ErrorHandler:

    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    
    Set RngShowHide = Nothing
End Sub

' ===============================================================
' GetDataAll
' gets all data from sheet for reporting
' ---------------------------------------------------------------
Public Function GetDataAll() As Variant()
    Dim AryDataAll() As Variant
    AryDataAll = Range("A4:AS" & GetCrewCountAll + 3)
    GetDataAll = AryDataAll
End Function

' ===============================================================
' GetPersDetails
' gets all personal details for export
' ---------------------------------------------------------------
Public Function GetPersDetails() As Variant()
    Dim AryPersDet As Variant
    AryPersDet = Range("A4:G" & GetCrewCountAll + 3)
    GetPersDetails = AryPersDet
End Function

' ===============================================================
' ClearPersDetails
' clears all personal details from sheet
' ---------------------------------------------------------------
Public Sub ClearPersDetails()
    Range("A4:G" & GetCrewCountAll + 3).ClearContents
End Sub

' ===============================================================
' WritePersDetails
' writes Personal Details from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Public Function WritePersDetails(ByRef AryPersDetails() As Variant) As Boolean
        
    On Error GoTo ErrorHandler
    
    Me.Unprotect SEC_KEY
    
    ClearPersDetails
    
    Range(RNG_PERS_DET & UBound(AryPersDetails) + 3) = AryPersDetails
    
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
       
    WritePersDetails = False
Exit Function

ErrorHandler:
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    WritePersDetails = True
End Function



