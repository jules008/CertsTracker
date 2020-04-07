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
' v1.0.1 - Changed reference to Intitialise routine
' v1.0.2 - formatting changes
' v1.0.3 - Use Enum for cols
' v1.0.4 - Show hide Cols
' v1.0.5 - Added global ranges
' v1.0.6 - Changed dev mode switch
' v1.1.2 - Changed refresh to array
' v1.1.3 - generic range changes and remove single line refresh
' v1.2.0 - Added function to return qual status
'---------------------------------------------------------------
' Date - 7 Apr 20
'===============================================================
Option Explicit

' ===============================================================
' Constants
' --------------------------------------------------------------
Private Const COL_WIDTHS As String = "20, 10, 20,10,15, 10, 10, 15, 10, 10"
Private Const NO_HEADERS As Integer = 3             'no of rows above data
Private Const RNG_PERS_DET As String = "A4:J"
Private Const NO_PERS_DET_COLS = 11
Private Const RNG_WORKING_START As String = "L"
Private Const RNG_WORKING_END As String = ":AX"
Private Const RNG_WORKING As String = "L4:AX"
Private Const RNG_PERSON_DET_START As String = "A"
Private Const RNG_PERSON_DET_END As String = "J"
Private Const RNG_DATA_ALL As String = "A4:AX"
Private Const RNG_ACTIVE As String = "J1:J1000"

' ===============================================================
' RefreshQuals
' Refreshes all qulaification data.
' ---------------------------------------------------------------
Public Sub RefreshQuals()
    Dim RngCell As Range
    Dim AryRoleQualArea() As Variant
    Dim RngRoleQualArea As Range
    Dim AryOutputQualArea() As Variant
    Dim RequiredYN As EnumTriState
    Dim GainedYN As EnumTriState
    Dim Role As EnumRole
    Dim Qual As EnumQual
    Dim Status As EnumStatus
    Dim QualStatus As EnumExpiryStatus
    Dim SSN As String
    Dim CourseDate As String
    Dim NoCrew As Integer
    Dim RowX, ColY, x, y As Integer
    
    On Error GoTo ErrorHandler
    
    NoCrew = ModMain.NoRows
    
    If NoCrew = 0 Then Exit Sub
    ReDim AryOutputQualArea(1 To NoCrew, 1 To NO_COURSES)
    
    Set RngRoleQualArea = Me.Range(RNG_WORKING & NoCrew + NO_HEADERS)
    
    AryRoleQualArea = RngRoleQualArea

    ModMain.PerfSettingsOn
    
    For x = LBound(AryRoleQualArea, 1) To UBound(AryRoleQualArea, 1)
        For y = LBound(AryRoleQualArea, 2) To UBound(AryRoleQualArea, 2)
        
            Role = GetRole(x + NO_HEADERS)
            Qual = GetQual(y + PERS_DET_NO_COLS + 1)
            SSN = GetSSN(x + NO_HEADERS)
            Status = GetStatus(x + NO_HEADERS)
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
                            If QualStatus = Valid Then AryOutputQualArea(x, y) = 4
                            If QualStatus = Expired Then AryOutputQualArea(x, y) = -5
                        End If
                        If Status = InActive Then AryOutputQualArea(x, y) = 9
                    Else
                        If Status = Active Then AryOutputQualArea(x, y) = 3
                        If Status = InActive Then AryOutputQualArea(x, y) = 8
                    End If
                Case Is = No
                    If GainedYN = Yes Then
                        If Status = Active Then
                            If QualStatus = Valid Then AryOutputQualArea(x, y) = 1
                            If QualStatus = Expired Then AryOutputQualArea(x, y) = -2
                        End If
                        If Status = InActive Then AryOutputQualArea(x, y) = 7
                    Else
                        If Status = Active Then AryOutputQualArea(x, y) = 0
                        If Status = InActive Then AryOutputQualArea(x, y) = 6
                    End If
                Case Is = Blank
                    AryOutputQualArea(x, y) = ""
            End Select
        Next y
    Next
    Range(RNG_WORKING & 1000).ClearContents
    RngRoleQualArea = AryOutputQualArea
    Set RngRoleQualArea = Nothing
    ModMain.PerfSettingsOff
Exit Sub

ErrorHandler:
    Set RngRoleQualArea = Nothing
    ModMain.PerfSettingsOff

End Sub

' ===============================================================
' GetSSN
' Returns SSN for active line
' ---------------------------------------------------------------
Private Function GetSSN(RowNo As Integer) As String
    GetSSN = Me.Cells(RowNo, eSSN)
End Function

' ===============================================================
' GetStatus
' Returns Status for active line
' ---------------------------------------------------------------
Private Function GetStatus(RowNo As Integer) As EnumStatus
    If Me.Cells(RowNo, eStatus) = "Active" Then GetStatus = Active
    If Me.Cells(RowNo, eStatus) = "Inactive" Then GetStatus = InActive
End Function

' ===============================================================
' GetRole
' Returns Role for active line
' ---------------------------------------------------------------
Private Function GetRole(ByVal RowNo As Integer) As EnumRole
    Dim StrRole As String
    
    StrRole = Me.Cells(RowNo, ePosition)
    
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
    Me.Cells(GetCrewCountAll + 4, eStatus) = "Active"
    Me.Cells(GetCrewCountAll + 4, eStatus).Select
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
' BtnShowHideCols_Click
' ---------------------------------------------------------------
Private Sub BtnShowHideCols_Click()
    FrmShowHideCols.Show
End Sub

' ===============================================================
' CmdRefresh_Click
' ---------------------------------------------------------------
Private Sub CmdRefresh_Click()
    Dim Prompt As Boolean
    
    If USER_LEVEL = DevLvl Then Prompt = True Else Prompt = False
    ModStartUp.InitialiseSystem Prompt
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
    
    If Not Intersect(Target, Me.Range(RNG_WORKING & CrewCnt + NO_HEADERS)) Is Nothing Then
        
        Debug.Print "FrmEnterDate.ShowForm", SSN, Qual
        FrmEnterDate.ShowForm SSN, Qual

        Me.RefreshQuals
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
    
    If Not Intersect(Target, Me.Range(RNG_PERSON_DET_START & NO_HEADERS & ":" & RNG_PERSON_DET_END & NoRows + NO_HEADERS)) Is Nothing Then
        Me.RefreshQuals
    End If
    
    If Not Intersect(Target, Me.Range(RNG_ACTIVE)) Is Nothing Then
        Me.RefreshQuals
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
    
    Set RngSSN = Me.Columns(eSSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then Exit Function
    
    GetName = RngResult.Offset(0, eName - eSSN)
    
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
    
    Set RngSort = Range("A4:AS" & ModMain.NoRows + NO_HEADERS)
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
    
    Set RngShowHide = Range("A" & NO_HEADERS & ":AS" & ModMain.NoRows + NO_HEADERS)
    
    Me.Unprotect SEC_KEY
    
    If CmdShowHide.Caption = "Hide Leavers" Then
        RngShowHide.AutoFilter Field:=eStatus, Criteria1:="Active", VisibleDropDown:=False
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
    AryDataAll = Range(RNG_DATA_ALL & GetCrewCountAll + NO_HEADERS)
    GetDataAll = AryDataAll
End Function

' ===============================================================
' GetPersDetails
' gets all personal details for export
' ---------------------------------------------------------------
Public Function GetPersDetails() As Variant()
    Dim AryPersDet As Variant
    AryPersDet = Range(RNG_PERS_DET & GetCrewCountAll + NO_HEADERS)
    GetPersDetails = AryPersDet
End Function

' ===============================================================
' ClearPersDetails
' clears all personal details from sheet
' ---------------------------------------------------------------
Public Sub ClearPersDetails()
    Range(RNG_PERS_DET & GetCrewCountAll + NO_HEADERS).ClearContents
End Sub

' ===============================================================
' ReadColStatus
' Returns whether the specified column is hidden
' ---------------------------------------------------------------
Public Function ReadColStatus(Col As EnumColsMain) As Boolean
    ReadColStatus = Not Me.Columns(Col).EntireColumn.Hidden
End Function

' ===============================================================
' SetColStatus
' hides or shows specified column
' ---------------------------------------------------------------
Public Sub SetColStatus(Col As EnumColsMain, Show As Boolean)
    Dim i As Integer
    Dim ColWidths() As String
    
    ColWidths = Split(COL_WIDTHS, ",")
    
    If Show Then
        Me.Columns(Col).EntireColumn.Hidden = False
        Me.Columns(Col).ColumnWidth = ColWidths(Col - 1)
    Else
        Me.Columns(Col).EntireColumn.Hidden = True
    End If
End Sub

' ===============================================================
' SetControlPos
' Sets the specified control's position
' ---------------------------------------------------------------
Public Sub SetControlPos(Cntrl As String, Top As Long, Left As Long, Height As Long, Width As Long)
    
    If Top + Left + Width + Height = 0 Then
        ShtMain.Shapes(Cntrl).Visible = False
    Else
        ShtMain.Shapes(Cntrl).Visible = True
        With ShtMain.Shapes(Cntrl)
            .Top = Top
            .Left = Left
            .Height = Height
            .Width = Width
        End With
    End If
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
    
    Range(RNG_PERS_DET & UBound(AryPersDetails) + NO_HEADERS) = AryPersDetails
    
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
       
    WritePersDetails = False
Exit Function

ErrorHandler:
    If USER_LEVEL <> DevLvl Then Me.Protect SEC_KEY
    WritePersDetails = True
End Function

' ===============================================================
' SecureCols
' Hides all sensitive columns.  used when closing or saving book
' ---------------------------------------------------------------
Public Sub SecureCols()
    
    SetColStatus eDoB, False
    SetColStatus eFINNo, False
    SetColStatus eDoDRef, False
    
    SetControlPos "CmoSortBy", 102, 252, 19, 80
    SetControlPos "CmdSort", 102, 330, 18, 33

End Sub

' ===============================================================
' GetQualStatus
' looks up the selected qualification status for the SSN provided
' ---------------------------------------------------------------
Public Function GetQualStatus(SSN As String, Qual As EnumQual) As Integer
    Dim RngSSN As Range
    Dim RngResult As Range
    Dim ResRowNo As Integer
    Dim ResColNo As Integer
    
    Set RngSSN = Me.Columns(eSSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then Exit Function
    ResRowNo = RngResult.Row
    
    'get col no from qual
    ResColNo = Qual + NO_PERS_DET_COLS
    
    GetQualStatus = Me.Cells(ResRowNo, ResColNo)
    
    Set RngResult = Nothing
    Set RngSSN = Nothing
End Function


