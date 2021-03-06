Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
' Start up functions
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - Added What's new Message
'---------------------------------------------------------------
' Date - 14 Apr 20
'===============================================================
Option Explicit

' ===============================================================
' InitialiseSystem
' Initialisation routine
' ---------------------------------------------------------------
Public Sub InitialiseSystem(Prompt As Boolean)
    PerfSettingsOn
    ModSecurity.DetectUser Prompt
    If ShtUsers.UpdateUserAccess Then WhatsNewMsg
    ShtMain.RefreshQuals
    PerfSettingsOff
End Sub

' ===============================================================
' WhatsNewMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub WhatsNewMsg()
        
    MsgBox " Updated to Version " & VERSION & " - What's New" _
                    & Chr(13) & " - New report of members requiring a selected certification" _
                    & Chr(13) & " - Fixed locked admin columns bug" _
                    & Chr(13) & " - Role data back up" _
                    & Chr(13) & "", vbOKOnly + vbInformation, "New version"



End Sub


