VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtCourseDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

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

Public Sub SaveCourseDate(SSN As String, Qual As EnumQual, CourseDate As Date)
    Dim RngSSN As Range
    Dim RngResult As Range
    
    Set RngSSN = Me.Range(RNG_SSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)


    Set RngResult = Nothing
    Set RngSSN = Nothing
End Sub

Public Function GetAllDates(SSN As String) As Variant()
    Dim RngSSN As Range
    Dim RngResult As Range
    Dim AryDates(1 To NO_COURSES) As Variant
    
    Set RngSSN = Me.Range(RNG_SSN)
    Set RngResult = RngSSN.Find(SSN, , xlValues, xlWhole, xlByRows, xlNext)
    
    If RngResult Is Nothing Then Exit Function
    
    Set RngResult = RngResult.Offset(0, 1).Resize(1, 37)
    
    GetAllDates = RngResult
    
    Set RngResult = Nothing
    Set RngSSN = Nothing

End Function

Public Function GetAllData() As Variant()
    Dim AryDataAll() As Variant
    AryDataAll = Range("B1:AN" & ShtMain.GetCrewCountAll + 3)
    GetAllData = AryDataAll

End Function

Public Sub ClearAllData()
    Range("B3:AN1000").ClearContents
End Sub

