Attribute VB_Name = "ModReports"

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
            AryQuals = ShtRoleLU.Range("B6:AL6")
            EligibleRnk = "Firefighter"
            ReportTitle = "Driver Operator Promotion Eligibility Report"
        
        Case DOtoCM
            Debug.Print DOtoCM
            AryQuals = ShtRoleLU.Range("B7:AL7")
            EligibleRnk = "Driver/Op"
            ReportTitle = "Crew Manager Promotion Eligibility Report"
        
        Case CMtoSC
            Debug.Print CMtoSC
            AryQuals = ShtRoleLU.Range("B8:AL8")
            EligibleRnk = "Crew Manager"
            ReportTitle = "Station Captain Promotion Eligibility Report"
        
        Case SCtoAC
            Debug.Print SCtoAC
            AryQuals = ShtRoleLU.Range("B12:AL12")
            EligibleRnk = "Station Captain"
            ReportTitle = "Assistant Chief Promotion Eligibility Report"
    
    End Select
    
    ArySource = ShtMain.GetDataAll
    
    'Loop through both arrays to look for qualifiation matches
    For i = LBound(ArySource) To UBound(ArySource)
    
        If ArySource(i, 3) = EligibleRnk Then
            
            Debug.Print ArySource(i, 1), "Eligible"
            Qualified = True
            
            For x = 1 To NO_COURSES
                
                If AryQuals(1, x) = 1 Then
                    If ArySource(i, x + 8) <> 1 And ArySource(i, x + 8) <> 4 Then Qualified = False
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

Public Sub ExpQualReport()
    Dim AryQuals() As Variant
    Dim AryReport(1 To 500, 1 To 5) As Variant
    Dim ArySource As Variant
    Dim SSN As String
    Dim Name As String
    Dim Qual As EnumQual
    Dim QDate As Date
    Dim Watch As String
    Dim Headings(0 To 5) As String
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
       
    ArySource = ShtMain.GetDataAll
    y = 1
    'Loop through array to look for expired qualifiations
    For i = LBound(ArySource) To UBound(ArySource)
        
        If ArySource(i, 7) = "Active" Then
            For x = 1 To NO_COURSES
                If ArySource(i, x + 8) < 0 Then
                    SSN = ArySource(i, 6)
                    Name = ArySource(i, 1)
                    Watch = ArySource(i, 5)
                    Qual = x
                    QDate = ShtCourseDates.LookUpCourseDate(SSN, Qual, eRead)
                    
                    AryReport(y, 1) = SSN
                    AryReport(y, 2) = Name
                    AryReport(y, 3) = Watch
                    AryReport(y, 4) = QualConvEnum(Qual)
                    AryReport(y, 5) = QDate
                    y = y + 1
                End If
            Next
        End If
    Next
    
    If y > 0 Then
        Headings(0) = "SSN"
        Headings(1) = "Name"
        Headings(2) = "Watch"
        Headings(3) = "Qualification"
        Headings(4) = "Date"

        ShtReport.PrintReport AryReport, "Expired Qualifications", Headings
    Else
        MsgBox "There were no results for the report", vbInformation + vbOKOnly
    End If
End Sub


