VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtCourseDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtCourseDates
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - Moved Write course dates into module
' v1.2.0 - Added function to return days to expiry
' v1.2.1 - Changed to generic ranges
'---------------------------------------------------------------
' Date - 23 Mar 20
'===============================================================
Option Explicit

Private Const RNG_COURSE_DATES As String = "B1:AO"

' ===============================================================
' LookUpCourseExp
' Checks the expiry date of the course if applicable.  Returns
' "" if not applicable
' ---------------------------------------------------------------
Public Function LookUpCourseExp(SSN As String, Qual As EnumQual) As String
    Dim QualNeeded As Boolean
    Dim RngSSN As Range
    Dim RngResult As Range
    Dim ResRowNo As Integer
    Dim ResColNo As Integer
    Dim QualDate As String
    Dim DaysToExp As Integer
    
    Set RngSSN = Me.Range(RNG_SSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then
        Set RngResult = RngSSN.End(xlDown)
        RngResult.Offset(1, 0) = SSN
    End If
    
    ResRowNo = RngResult.Row
    ResColNo = Qual + 2
    
    QualDate = Me.Cells(ResRowNo, ResColNo)
    
    If QualDate <> "" Then
        LookUpCourseExp = ShtRoleLU.DaysTillExpiry(QualDate, Qual)
    Else
        LookUpCourseExp = 999
    End If
    
    Set RngResult = Nothing
    Set RngSSN = Nothing
End Function

' ===============================================================
' LookUpCourseDate
' Looks to see if a course has been completed and returns date.
' Also can write the date on update
' ---------------------------------------------------------------
Public Function LookUpCourseDate(SSN As String, Qual As EnumQual, ReadWrite As EnumRW, Optional CourseDate As String) As String
    Dim QualNeeded As Boolean
    Dim RngSSN As Range
    Dim RngResult As Range
    Dim ResRowNo As Integer
    Dim ResColNo As Integer
    Dim QualDate As String
    
    Set RngSSN = Me.Range(RNG_SSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    'get row no from ssn
    If RngResult Is Nothing Then
        Set RngResult = RngSSN.End(xlDown)
        RngResult.Offset(1, 0) = SSN
    End If
    
    ResRowNo = RngResult.Row
    
    'get col no from qual
    ResColNo = Qual + 2
    
    Select Case ReadWrite
        Case Is = eRead
            QualDate = Me.Cells(ResRowNo, ResColNo)
'            If QualDate = "1" Then QualDate = "Passed"
            LookUpCourseDate = QualDate
        
        Case Is = EWrite
            Me.Cells(ResRowNo, ResColNo) = CourseDate
            
        Case Is = EClear
            Me.Cells(ResRowNo, ResColNo) = ""
            
    End Select
    
    Set RngResult = Nothing
    Set RngSSN = Nothing
End Function

' ===============================================================
' GetAllDates
' Returns all dates from SSN
' ---------------------------------------------------------------
Public Function GetAllDates(SSN As String) As Variant()
    Dim RngSSN As Range
    Dim RngResult As Range
    Dim AryDates(1 To NO_COURSES) As Variant
    
    Set RngSSN = Me.Range(RNG_SSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then Exit Function
    
    Set RngResult = RngResult.Offset(0, 1).Resize(1, NO_COURSES)
    
    GetAllDates = RngResult
    
    Set RngResult = Nothing
    Set RngSSN = Nothing

End Function

' ===============================================================
' GetAllData
' Returns all data from sheet
' ---------------------------------------------------------------
Public Function GetAllData() As Variant()
    Dim AryDataAll() As Variant
    AryDataAll = Range(RNG_COURSE_DATES & ShtMain.GetCrewCountAll + 3)
    GetAllData = AryDataAll

End Function

' ===============================================================
' ClearAllData
' Clears all data from sheet
' ---------------------------------------------------------------
Public Sub ClearAllData()
    Range(RNG_COURSE_DATES & "1000").ClearContents
End Sub

' ===============================================================
' WriteCourseDates
' writes Course Dates from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Public Function WriteCourseDates(ByRef AryCourseDates() As Variant) As Boolean
        
    On Error GoTo ErrorHandler
    
    ShtCourseDates.Unprotect SEC_KEY
    
    ClearAllData
    
    Range(RNG_COURSE_DATES & UBound(AryCourseDates)) = AryCourseDates
       
    If USER_LEVEL <> DevLvl Then ShtCourseDates.Protect SEC_KEY
       
    WriteCourseDates = False
Exit Function

ErrorHandler:
    If USER_LEVEL <> DevLvl Then ShtCourseDates.Protect SEC_KEY
    WriteCourseDates = True
End Function


