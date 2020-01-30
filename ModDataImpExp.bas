Attribute VB_Name = "ModDataImpExp"
Private Function ExportPersDet(FilePath As String) As Boolean
    Dim AryPersDet() As Variant
    Dim CrewCount As Integer
    Dim Rw As Integer
    Dim Cl As Integer
    Dim TxtLine As String
    
    On Error GoTo ErrorHandler
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set ExpFile = FSO.CreateTextFile(FilePath & "\UserDetails.txt", True)
    
    AryPersDet = ShtMain.GetPersDetails
    
    For Rw = LBound(AryPersDet) To UBound(AryPersDet)
        TxtLine = ""
        For Cl = 1 To 7
            TxtLine = TxtLine & AryPersDet(Rw, Cl) & ";"
        Next
        ExpFile.WriteLine (TxtLine)
    Next
    ExpFile.Close
    
    ExportPersDet = False
    
    Set ExpFile = Nothing
    Set FSO = Nothing
Exit Function

ErrorHandler:
    Set ExpFile = Nothing
    Set FSO = Nothing
    ExportPersDet = True
End Function

Private Function ExportCourseDates(FilePath As String) As Boolean
    Dim AryDates() As Variant
    Dim CrewCount As Integer
    Dim Rw As Integer
    Dim Cl As Integer
    Dim TxtLine As String
    
    On errror GoTo ErrorHandler
    
    AryDates = ShtCourseDates.GetAllData
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ExpFile = FSO.CreateTextFile(FilePath & "\CourseDates.txt", True)
    
    For Rw = LBound(AryDates) To UBound(AryDates)
        TxtLine = ""
        For Cl = 1 To 38
            TxtLine = TxtLine & AryDates(Rw, Cl) & ";"
        Next
        ExpFile.WriteLine (TxtLine)
    Next
    ExpFile.Close
    
    ExportCourseDates = False
    
    Set ExpFile = Nothing
    Set FSO = Nothing
Exit Function

ErrorHandler:
    Set ExpFile = Nothing
    Set FSO = Nothing
    ExportCourseDates = True
End Function

Public Sub ExportData()
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim ErrFlag1 As Boolean
    Dim ErrFlag2 As Boolean
    
    On errror GoTo ErrorHandler
    
    Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Fldr
        .Title = "Select Destination"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    ErrFlag1 = ModDataImpExp.ExportCourseDates(FilePath)
    ErrFlag2 = ModDataImpExp.ExportPersDet(FilePath)
    
    If ErrFlag1 Or ErrFlag2 Then GoTo ErrorHandler
    
    MsgBox "Export Complete", vbOKOnly + vbInformation, "Data Export"
    
    Set Fldr = Nothing

Exit Sub

ErrorHandler:
    Set Fldr = Nothing
    
        MsgBox "An error with the export has occured", vbOKOnly + vbCritical, "Error"

End Sub

Private Sub ClearAllData()
    ShtMain.Unprotect "2683174"
    
    ShtMain.AutoFilterMode = False
    ShtMain.CmdShowHide.Caption = "Hide Leavers"
    ShtMain.ClearPersDetails
    ShtCourseDates.ClearAllData
    
    ShtMain.Protect "2683174"

End Sub

Public Sub ImportData()
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim ErrFlag1 As Boolean
    Dim ErrFlag2 As Boolean
    
    On Error GoTo ErrorHandler
    
    Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Fldr
        .Title = "Select Input Files Location"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    ErrFlag1 = ImportPersData(FilePath)
    ErrFlag2 = ImportCourseDates(FilePath)
    
    If ErrFlag1 Or ErrFlag2 Then GoTo ErrorHandler
    
    MsgBox "Export Complete", vbOKOnly + vbInformation, "Data Export"
    
    Set Fldr = Nothing

Exit Sub

ErrorHandler:
    Set Fldr = Nothing
    
        MsgBox "An error with the export has occured", vbOKOnly + vbCritical, "Error"

End Sub

Private Function ImportPersData(FilePath) As Boolean
    Dim AryImport() As String
    Dim FullFilePath As String
    Dim TotalLines As Integer
    
    On Error GoTo ErrorHandler
    
    FullFilePath = FilePath & "UserDetails.txt"
    
    AryImport = DelimitedTextFileToArray(FullFilePath)
    TotalLines = UBound(AryImport)
    
    ShtMain.Range("A4:G" & TotalLines + 3) = AryImport
    
    ImportPersData = False
Exit Function

ErrorHandler:
    ImportPersData = True
End Function

Private Function DelimitedTextFileToArray(FilePath As String) As String()
    Dim Delimiter As String
    Dim TextFile As Integer
    Dim FileContent As String
    Dim LineArray() As String
    Dim DataArray() As String
    Dim TempArray() As String
    Dim Rows As Integer
    Dim rw As Long, col As Long
    
    Delimiter = ";"
    rw = 0
    
    TextFile = FreeFile
    Open FilePath For Input As TextFile
    
    FileContent = Input(LOF(TextFile), TextFile)
    
    Close TextFile
    
    LineArray() = Split(FileContent, vbCrLf)
    Rows = UBound(LineArray)
    ReDim DataArray(Rows, 39)
    
    For x = LBound(LineArray) To UBound(LineArray)
        If Len(Trim(LineArray(x))) <> 0 Then
            TempArray = Split(LineArray(x), Delimiter)
            
            For y = LBound(TempArray) To UBound(TempArray)
                DataArray(rw, y) = TempArray(y)
            Next y
        End If
    
        rw = rw + 1
    Next x
    DelimitedTextFileToArray = DataArray()
End Function

Private Function ImportCourseDates(FilePath) As Boolean
    Dim AryImport() As String
    Dim FullFilePath As String
    Dim TotalLines As Integer
    
    On Error GoTo ErrorHandler
    
    FullFilePath = FilePath & "CourseDates.txt"
    
    AryImport = DelimitedTextFileToArray(FullFilePath)
    TotalLines = UBound(AryImport)
    
    ShtCourseDates.Range("B1:AN" & TotalLines + 3) = AryImport
    
    ImportCourseDates = False
Exit Function

ErrorHandler:
    ImportCourseDates = True
End Function

