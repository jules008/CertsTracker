Attribute VB_Name = "Template_File_Handling"
Option Explicit

' File Dialog to select Folder
' ============================
Function GetFolder() As String
    Dim Fldr As FileDialog
    Dim FilePath As String
    Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With Fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo ExitDialog
        FilePath = .SelectedItems(1)
    End With
    
ExitDialog:
    GetFolder = FilePath
    Set Fldr = Nothing
End Function

' Create and fill text file
' ============================
' Enable reference to Microsoft Scripting Runtime

Sub TextFile()
    Dim TxtLine As String
    Dim FSO As FileSystemObject
    Dim ExpFile As TextStream
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ExpFile = FSO.CreateTextFile("e:\UserDetails.txt", True)
    
    TxtLine = "Line of text"
    ExpFile.WriteLine (TxtLine)
    ExpFile.Close
End Sub

