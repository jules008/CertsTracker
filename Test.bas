Attribute VB_Name = "Test"

Public Sub TestQual()
    Dim Qual As EnumQual
    Dim Role As EnumRole
    
    Qual = HazmatTech
    Role = Firefighter
    
End Sub

Public Sub TestCourseDate()
    Dim Qual As EnumQual
    Dim SSN As String
    
    SSN = "405-44-8124"
    Qual = HazmatAW
    
End Sub


Public Sub Refresh()
    ShtMain.RefreshQuals
End Sub

Public Sub DeleteZeros()
    Dim Cell As Range
    
    For Each Cell In Selection
        If Cell = 0 Then Cell = ""
    
    
    Next
End Sub

Public Sub TestGetAlldays()
    Dim SSN As String
    Dim AryDates() As Variant
    
    SSN = "405-44-0342"
    
    FrmTngDates.ShowForm (SSN)
End Sub

Public Sub FontSize()
    Dim i As Integer
    
    With FrmTngDates
        For i = 1 To 37
            
            .Controls("TxtDates" & i).Value = "01 Dec 19"
            .Controls("TxtDates" & i).Font = "Trebuchet MS"
            .Controls("TxtDates" & i).Font.Size = 10
        Next
        
        .Show
        
    End With
End Sub

Public Sub Charts()
    Dim ChartTrend As Chart
    Set ChartTrend = ShtDashboard.ChartObjects(1).Chart
    
    
    For i = 1 To 2
        Debug.Print ShtDashboard.ChartObjects(i).Select
    Next
    
    
End Sub

Public Sub SortBy()
    ShtMain.SortBy 4
End Sub

Public Sub TestExpiry()
    Dim Status As EnumExpiryStatus
    
    Status = ShtRoleLU.RetQualStatus("13 dec 17", FirefighterI)
    Debug.Print Status
End Sub

Public Sub TestLogin()
    Dim UserLog As EnumUserLvl
    
    UserLog = FrmLogin.ShowForm
    
    Debug.Print UserLog
End Sub

Public Sub testArray()
    Dim AryTest() As Variant
    
    AryTest = ShtMain.GetDataAll
End Sub
