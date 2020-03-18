Attribute VB_Name = "ModSecurity"
'===============================================================
' Module ModSecurity
'===============================================================
' v1.0.0 - Initial Version
' v1.0.1 - added new ShtUsers and moved user level detection
' v1.0.2 - Changes to Password
'---------------------------------------------------------------
' Date - 17 Mar 20
'===============================================================
Option Explicit

Public Sub BasicView()
    
    'Sheet
    ShtMain.Unprotect SEC_KEY
    ShtMain.Range("A:H").Locked = True
    ShtMain.CmdReports.Visible = False
    ShtMain.BtnImpExp.Visible = False
    ShtMain.BtnAddNew.Visible = False
    ShtMain.Shapes("TxtView").Visible = msoFalse
    
    ShtReport.Visible = xlSheetVeryHidden
    ShtUsers.Visible = xlSheetVeryHidden
    ShtColours.Visible = xlSheetVeryHidden
    ShtCourseDates.Visible = xlSheetVeryHidden
    ShtLists.Visible = xlSheetVeryHidden
    ShtRoleLU.Visible = xlSheetVeryHidden
    
    ShtDashboard.Protect SEC_KEY
    ShtMain.Protect SEC_KEY
    
    USER_LEVEL = BasicLvl
End Sub

Public Sub AdminView()
    ShtMain.Unprotect SEC_KEY
    ShtMain.Range("A:G").Locked = False
    ShtMain.CmdReports.Visible = True
    ShtMain.BtnImpExp.Visible = False
    ShtMain.BtnAddNew.Visible = True
    ShtMain.Shapes("TxtView").Visible = msoCTrue
    ShtMain.Shapes("TxtView").TextFrame.Characters.Text = "Administrator View"
    
    ShtReport.Visible = xlSheetVeryHidden
    ShtUsers.Visible = xlSheetVeryHidden
    ShtColours.Visible = xlSheetVeryHidden
    ShtCourseDates.Visible = xlSheetVeryHidden
    ShtLists.Visible = xlSheetHidden
    ShtRoleLU.Visible = xlSheetHidden
     
    ShtDashboard.Protect SEC_KEY
    ShtMain.Protect SEC_KEY
       
    USER_LEVEL = AdminLvl
End Sub

Public Sub DevView()
    ShtMain.Unprotect SEC_KEY
    ShtMain.Range("A:G").Locked = False
    ShtMain.CmdReports.Visible = True
    ShtMain.BtnImpExp.Visible = True
    ShtMain.BtnAddNew.Visible = True
    ShtMain.Shapes("TxtView").Visible = msoTrue
    ShtMain.Shapes("TxtView").TextFrame.Characters.Text = "Developer View"
    
    ShtReport.Visible = xlSheetVisible
    ShtUsers.Visible = xlSheetVisible
    ShtColours.Visible = xlSheetVisible
    ShtCourseDates.Visible = xlSheetVisible
    ShtLists.Visible = xlSheetVisible
    ShtRoleLU.Visible = xlSheetVisible
    
    ShtDashboard.Unprotect SEC_KEY
    
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
    
    If USER_LEVEL = DevLvl Then USER_LEVEL = FrmLogin.ShowForm
    If USER_LEVEL = DevLvl Then DevView
    If USER_LEVEL = AdminLvl Then AdminView
    If USER_LEVEL = BasicLvl Then BasicView
        
End Sub
