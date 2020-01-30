Attribute VB_Name = "ModSecurity"

Public Sub BasicView()
    ShtMain.Unprotect "2683174"
    ShtMain.Range("A:H").Locked = True
    ShtReport.Visible = xlSheetVeryHidden
    ShtColours.Visible = xlSheetVeryHidden
    ShtCourseDates.Visible = xlSheetVeryHidden
    ShtLists.Visible = xlSheetVeryHidden
    ShtRoleLU.Visible = xlSheetVeryHidden
    ShtMain.CmdReports.Visible = False
    ShtMain.BtnImpExp.Visible = False
    
    ShtMain.Shapes("TxtView").Visible = msoFalse
    ShtDashboard.Protect "2683174"
    ShtMain.Protect "2683174"
    
    USER_LEVEL = BasicLvl
End Sub

Public Sub AdminView()
    ShtDashboard.Protect "2683174"
    ShtMain.Unprotect "2683174"
    ShtMain.Range("A:G").Locked = False
    ShtReport.Visible = xlSheetVeryHidden
    ShtColours.Visible = xlSheetVeryHidden
    ShtCourseDates.Visible = xlSheetVeryHidden
    ShtLists.Visible = xlSheetHidden
    ShtRoleLU.Visible = xlSheetHidden
    
    ShtMain.Shapes("TxtView").Visible = msoCTrue
    ShtMain.Shapes("TxtView").TextFrame.Characters.Text = "Administrator View"
    ShtMain.CmdReports.Visible = True
    ShtMain.BtnImpExp.Visible = False
    ShtMain.Protect "2683174"
    
    USER_LEVEL = AdminLvl
End Sub

Public Sub DevView()
    ShtMain.Unprotect "2683174"
    ShtDashboard.Unprotect "2683174"
    ShtMain.Range("A:G").Locked = False
    ShtReport.Visible = xlSheetVisible
    ShtColours.Visible = xlSheetVisible
    ShtCourseDates.Visible = xlSheetVisible
    ShtLists.Visible = xlSheetVisible
    ShtRoleLU.Visible = xlSheetVisible
    
    
    ShtMain.Shapes("TxtView").Visible = msoTrue
    ShtMain.Shapes("TxtView").TextFrame.Characters.Text = "Developer View"
    ShtMain.CmdReports.Visible = True
    ShtMain.BtnImpExp.Visible = True
    
    USER_LEVEL = DevLvl
End Sub


Private Sub MenuOff()
    With Application
    End With
End Sub

Public Sub DetectUser()
    If Application.UserName = "MOHR, CHRISTOPHER M GS-10 USAF USAFE 423 CES/CEF" Then
        ModGlobals.USER_LEVEL = AdminLvl
    Else
        If Application.UserName = "TURNER, JULIAN D GB USAF USAFE 423 CES/CEF" Or Application.UserName = "Julian Turner" Then
            ModGlobals.USER_LEVEL = DevLvl
        Else
            ModGlobals.USER_LEVEL = BasicLvl
        End If
    End If
        
End Sub
