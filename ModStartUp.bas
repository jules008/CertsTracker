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
        
    MsgBox "Version " & VERSION & " - What's New" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Added USAR as a Station" _
                    & Chr(13) & "", vbOKCancel + vbInformation, "New version"



End Sub


