Attribute VB_Name = "ModReports"
'===============================================================
' Module ModReports
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - Improved Reporting
' v1.1.1 - added global ranges
' v1.2.0 - Required Qualification report
'---------------------------------------------------------------
' Date - 7 Apr 20
'===============================================================
Option Explicit

'===============================================================
' PromReports
' Reports on all members that are qualified for promotion to the
' next rank
'---------------------------------------------------------------
Public Sub PromReports(Report As EnumReport)
    Dim AryQuals() As Variant
    Dim AryReport(1 To 50, 1 To 5) As Variant
    Dim ArySource As Variant
    Dim EligibleRnk As String
    Dim Qualified As Boolean
    Dim ReportTitle As String
    Dim Headings(0 To 5) As String
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    Select Case Report
        Case FFtoDO
            Debug.Print FFtoDO
            AryQuals = ShtRoleLU.GetRoleEligibility(DriverOp)
            EligibleRnk = "Firefighter"
            ReportTitle = "Driver Operator Promotion Eligibility Report"
        
        Case DOtoCM
            Debug.Print DOtoCM
            AryQuals = ShtRoleLU.GetRoleEligibility(CrewManager)
            EligibleRnk = "Driver/Op"
            ReportTitle = "Crew Manager Promotion Eligibility Report"
        
        Case CMtoSC
            Debug.Print CMtoSC
            AryQuals = ShtRoleLU.GetRoleEligibility(StationCaptain)
            EligibleRnk = "Crew Manager"
            ReportTitle = "Station Captain Promotion Eligibility Report"
        
        Case SCtoAC
            Debug.Print SCtoAC
            AryQuals = ShtRoleLU.GetRoleEligibility(ACOps)
            EligibleRnk = "Station Captain"
            ReportTitle = "Assistant Chief Promotion Eligibility Report"
    
    End Select
    
    ArySource = ShtMain.GetDataAll
    
    'Loop through both arrays to look for qualifiation matches
    For i = LBound(ArySource) To UBound(ArySource)
    
        If ArySource(i, ePosition) = EligibleRnk Then
            
            Debug.Print ArySource(i, 1), "Eligible"
            Qualified = True
            
            For x = 1 To NO_COURSES
                
                If AryQuals(1, x) = 1 Then
                    
                    If ArySource(i, x + PERS_DET_NO_COLS + 1) <> 1 And ArySource(i, x + PERS_DET_NO_COLS + 1) <> 4 Then Qualified = False
                End If
            Next
            
            If Qualified Then
                y = y + 1
                AryReport(y, 1) = ArySource(i, 6)
                AryReport(y, 2) = ArySource(i, 1)
                AryReport(y, 3) = ArySource(i, 3)
                AryReport(y, 4) = ArySource(i, 4)
                AryReport(y, 5) = ArySource(i, 5)
            End If
        End If
    Next
    
    If y > 0 Then
        Headings(0) = "SSN"
        Headings(1) = "Name"
        Headings(2) = "Role"
        Headings(3) = "Contract"
        Headings(4) = "Watch"
        
        ShtReport.PrintReport AryReport, ReportTitle, Headings
    Else
        MsgBox "There were no results for the report", vbInformation + vbOKOnly
    End If
End Sub

'===============================================================
' ReqQualReport
' Reports on members that require the selected qualification
'---------------------------------------------------------------
Public Sub ReqQualReport(SelWatch As String, Qual As EnumQual)
    Dim AryQuals() As Variant
    Dim AryReport(1 To 500, 1 To 5) As Variant
    Dim ArySource As Variant
    Dim SSN As String
    Dim Name As String
    Dim Status As String
    Dim Position As String
    Dim Grade As String
    Dim QStatus As Integer
    Dim QDate As String
    Dim Watch As String
    Dim DaysToExp As Single
    Dim Headings(0 To 5) As String
    Dim Title As String
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
       
    ArySource = ShtMain.GetDataAll
    
    For i = LBound(ArySource) To UBound(ArySource)
        SSN = ArySource(i, eSSN)
        Name = ArySource(i, eName)
        Watch = ArySource(i, eWatch)
        Status = ArySource(i, eStatus)
        Position = ArySource(i, ePosition)
        Grade = ArySource(i, eGrade)
        
        Debug.Print SSN, Name, Watch, Status
        
        If Status = "Active" And _
            (Watch = SelWatch Or _
            SelWatch = "All") Then

            QStatus = ShtMain.GetQualStatus(SSN, Qual)
            
            If QStatus = 3 Then
                y = y + 1
                AryReport(y, 1) = Name
                AryReport(y, 2) = Grade
                AryReport(y, 3) = Position
                AryReport(y, 4) = Watch
                AryReport(y, 5) = QualConvEnum(Qual)
            End If
        End If
    Next
    
    If y > 0 Then
        Headings(0) = "Name"
        Headings(1) = "Grade"
        Headings(2) = "Position"
        Headings(3) = "Watch"
        Headings(4) = "Req Qual"
        
        Title = "Personnel Requiring " & QualConvEnum(Qual) & " Certification"
        
        ShtReport.PrintReport AryReport, Title, Headings
    Else
        MsgBox "There were no results for the report", vbInformation + vbOKOnly
    End If
End Sub

'===============================================================
' ExpQualReport
' Reports on expired qualifications
'---------------------------------------------------------------
Public Sub ExpQualReport(SelWatch As String, Period As Integer)
    Dim AryQuals() As Variant
    Dim AryReport(1 To 500, 1 To 5) As Variant
    Dim ArySource As Variant
    Dim SSN As String
    Dim Name As String
    Dim Status As String
    Dim Qual As EnumQual
    Dim QDate As String
    Dim Watch As String
    Dim DaysToExp As Integer
    Dim Headings(0 To 5) As String
    Dim Title As String
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
       
    ArySource = ShtMain.GetDataAll
    y = 0
    
    For i = LBound(ArySource) To UBound(ArySource)
        SSN = ArySource(i, eSSN)
        Name = ArySource(i, eName)
        Watch = ArySource(i, eWatch)
        Status = ArySource(i, eStatus)
        
        Debug.Print SSN, Name, Watch, Status
        
        If Status = "Active" And _
            (Watch = SelWatch Or _
            SelWatch = "All") Then

            For x = 1 To NO_COURSES
                Qual = x
                QDate = ShtCourseDates.LookUpCourseDate(SSN, Qual, eRead)
                DaysToExp = ShtCourseDates.LookUpCourseExp(SSN, Qual)
                If DaysToExp < Period Then
                    y = y + 1
                    AryReport(y, 1) = Name
                    AryReport(y, 2) = Watch
                    AryReport(y, 3) = QualConvEnum(Qual)
                    AryReport(y, 4) = QDate
                    AryReport(y, 5) = DaysToExp
                End If
            Next
        End If
    Next
    
    If y > 0 Then
        Headings(0) = "Name"
        Headings(1) = "Watch"
        Headings(2) = "Qualification"
        Headings(3) = "Date"
        Headings(4) = "Days Till Exp"
        
        Select Case Period
            Case 0
                Title = "Expired Qualifications"
            Case 10
                Title = "Qualifications Due Within 10 Days"
            Case 30
                Title = "Qualifications Due Within 30 Days"
            Case 60
                Title = "Qualifications Due Within 60 Days"
            End Select
        ShtReport.PrintReport AryReport, Title, Headings
    Else
        MsgBox "There were no results for the report", vbInformation + vbOKOnly
    End If
End Sub

