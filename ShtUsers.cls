VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtUsers
' Manages the list of users that have accessed the system.  Adds
' Names when logged in and then records if the 'what's new' message
' has been seen by the user.  On version change, the flag is cleared.
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 14 Mar 20
'===============================================================
Option Explicit
' ===============================================================
' Global Declarations
' ---------------------------------------------------------------
Private Const NO_COLS As Integer = 3
Private Const COL_USER As Integer = 1
Private Const COL_LAST_USED As Integer = 2
Private Const COL_SHOW_MSG As Integer = 3
' ===============================================================
' UpdateUserAccess
' Reads username from system and then, if it is not in the list, calls
' a routine to add it.  Updates the last access date with the current date.
' Returns true or false whether a whats new message is needed after update
' ---------------------------------------------------------------
Public Function UpdateUserAccess() As Boolean
    Dim RngWorking As Range
    Dim UserName As String
    Dim RngTgt As Range
    Dim NoUsers As Integer
    
    NoUsers = [no_users]
    
    UserName = Application.UserName
    
    If NoUsers = 0 Then
        Set RngTgt = Me.Cells(2, COL_USER)
        RngTgt = UserName
        RngTgt.Offset(0, 1) = Format(Now, "dd mmm yy")
        RngTgt.Offset(0, 2) = False
        UpdateUserAccess = True
    Else
        Set RngWorking = Me.Range(Me.Cells(2, 1), Me.Cells(NoUsers + 1, NO_COLS))
                
        Set RngTgt = RngWorking.Find(UserName)
            
        If RngTgt Is Nothing Then
            Set RngTgt = Range("A1").End(xlDown).Offset(1)
            RngTgt = UserName
            RngTgt.Offset(0, 1) = Format(Now, "dd mmm yy")
            RngTgt.Offset(0, 2) = False
            UpdateUserAccess = True
        Else
            RngTgt.Offset(0, 1) = Format(Now, "dd mmm yy")
            If RngTgt.Offset(0, 2) = True Then
                RngTgt.Offset(0, 2) = False
                UpdateUserAccess = True
            Else
                RngTgt.Offset(0, 2) = False
            End If
        End If
    End If
    
    Set RngWorking = Nothing
    Set RngTgt = Nothing
End Function

' ===============================================================
' ClearMsgFlag
' Clears the flag for all users on version update.
' ---------------------------------------------------------------
Public Sub ClearMsgFlag()
    If [no_users] > 0 Then Me.Cells(2, COL_SHOW_MSG) = True
    If [no_users] > 1 Then
        Range(Cells(2, COL_SHOW_MSG), Cells([no_users] + 1, COL_SHOW_MSG)).FillDown
    End If
End Sub

