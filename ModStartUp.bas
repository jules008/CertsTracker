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
' Date - 14 Mar 20
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
        
    MsgBox " Updated to Version " & VERSION & " - What's New" _
                    & Chr(13) & "" _
                    & Chr(13) & " - This What's New Box!" _
                    & Chr(13) & " - Improved Import and Export routine" _
                    & Chr(13) & " - Fixed Trend Issue" _
                    & Chr(13) & " - General Bug Fixes" _
                    & Chr(13) & " - More to follow....." _
                    & Chr(13) & "", vbOKOnly + vbInformation, "New version"



End Sub


