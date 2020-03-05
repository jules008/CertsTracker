Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
' Start up functions
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 05 Mar 20
'===============================================================
Option Explicit

' ===============================================================
' InitialiseSystem
' Initialisation routine
' ---------------------------------------------------------------
Public Sub InitialiseSystem()
    PerfSettingsOn
    ModSecurity.DetectUser
    If ShtUsers.UpdateUserAccess Then WhatsNewMsg
    ShtMain.RefreshQuals
    PerfSettingsOff
End Sub

' ===============================================================
' WhatsNewMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub WhatsNewMsg()
        
'        .Fields("SystemMessage") = "Version " & VERSION & " - What's New" _
'                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
'                    & Chr(13) & "" _
'                    & Chr(13) & " - Added USAR as a Station" _
'                    & Chr(13) & ""
'
'        .Fields("ReleaseNotes") = "Software Version: " & VERSION _
'                    & Chr(13) & "Database Version: " & DB_VER _
'                    & Chr(13) & "Date: " & VER_DATE _
'                    & Chr(13) & "" _
'                    & Chr(13) & "-  Added USAR as a Station - Items can now be ordered and allocated to USAR" _
'                    & Chr(13) & ""
'        .Update
'    End With
'
'    'reset read flags
'    Db.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
'
'    Set RstMessage = Nothing
'
End Sub


