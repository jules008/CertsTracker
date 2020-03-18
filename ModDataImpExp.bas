Attribute VB_Name = "ModDataImpExp"
'===============================================================
' Module ModDataImpExp
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - Added Trend data archive
'---------------------------------------------------------------
' Date - 17 Mar 20
'===============================================================
Option Explicit

Dim FSO As Scripting.FileSystemObject
' ===============================================================
' ExportPersDet
' Exports personal data to a text file and saves at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ExportPersDet(FilePath As String) As Boolean
    Dim AryPersDet() As Variant
    Dim CrewCount As Integer
    Dim rw As Integer
    Dim Cl As Integer
    Dim TxtLine As String
    Dim ExpFile As TextStream
    
    On Error GoTo ErrorHandler
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set ExpFile = FSO.CreateTextFile(FilePath & "\UserDetails.txt", True)
    
    AryPersDet = ShtMain.GetPersDetails
    
    For rw = LBound(AryPersDet) To UBound(AryPersDet)
        TxtLine = ""
        For Cl = 1 To 7
            TxtLine = TxtLine & AryPersDet(rw, Cl) & ";"
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

' ===============================================================
' ExportCourseDates
' Exports Course Dates to a text file and saves at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ExportCourseDates(FilePath As String) As Boolean
    Dim AryDates() As Variant
    Dim CrewCount As Integer
    Dim rw As Integer
    Dim Cl As Integer
    Dim TxtLine As String
    Dim ExpFile As TextStream
    
    On Error GoTo ErrorHandler
    
    AryDates = ShtCourseDates.GetAllData
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ExpFile = FSO.CreateTextFile(FilePath & "\CourseDates.txt", True)
    
    For rw = LBound(AryDates) To UBound(AryDates)
        TxtLine = ""
        For Cl = 1 To 38
            TxtLine = TxtLine & AryDates(rw, Cl) & ";"
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

' ===============================================================
' ExportData
' Main routine to carry out data export.  Calls other routines to
' do the actual export.
' ---------------------------------------------------------------
Public Sub ExportData()
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim ErrFlag1 As Boolean
    Dim ErrFlag2 As Boolean
    Dim ErrFlag3 As Boolean
    
    On Error GoTo ErrorHandler
    
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
    ErrFlag3 = ModDataImpExp.ExportTrendData(FilePath)
    
    If ErrFlag1 Or ErrFlag2 Or ErrFlag3 Then GoTo ErrorHandler
    
    MsgBox "Export Complete", vbOKOnly + vbInformation, "Data Export"
    
    Set Fldr = Nothing

Exit Sub

ErrorHandler:
    Set Fldr = Nothing
    
        MsgBox "An error with the export has occured", vbOKOnly + vbCritical, "Error"

End Sub

' ===============================================================
' ClearAllData
' Clears all data from sheet.
' ---------------------------------------------------------------
Public Sub ClearAllData()
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to clear all datails!!", vbCritical + vbYesNo, "Clear Datail")
    
    If Response = 6 Then
        ShtMain.Unprotect SEC_KEY
        ShtCourseDates.Unprotect SEC_KEY
        ShtDashboard.Unprotect SEC_KEY
        
        ShtMain.AutoFilterMode = False
        ShtMain.CmdShowHide.Caption = "Hide Leavers"
        ShtMain.ClearPersDetails
        ShtCourseDates.ClearAllData
        
        If USER_LEVEL <> DevLvl Then ShtMain.Protect SEC_KEY
        If USER_LEVEL <> DevLvl Then ShtCourseDates.Protect SEC_KEY
        If USER_LEVEL <> DevLvl Then ShtDashboard.Protect SEC_KEY
    End If
End Sub

' ===============================================================
' ImportData
' Main routine to import data from text files.  Calls other routines
' to do the actual data import.
' ---------------------------------------------------------------
Public Sub ImportData()
    Dim Fldr As FileDialog
    Dim FilePath As String
    Dim ErrFlag1 As Boolean
    Dim ErrFlag2 As Boolean
    Dim ErrFlag3 As Boolean
    
    On Error GoTo ErrorHandler
    
    ClearAllData
    
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
    ErrFlag3 = ImportTrendData(FilePath)
    
    If ErrFlag1 Or ErrFlag2 Or ErrFlag3 Then GoTo ErrorHandler
    
    MsgBox "Import Complete", vbOKOnly + vbInformation, "Data Export"
    
    Set Fldr = Nothing

Exit Sub

ErrorHandler:
    Set Fldr = Nothing
    
        MsgBox "An error with the export has occured", vbOKOnly + vbCritical, "Error"

End Sub

' ===============================================================
' ImportPersData
' Imports Personal Data from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ImportPersData(FilePath) As Boolean
    Dim AryImport() As Variant
    Dim FullFilePath As String
    Dim TotalLines As Integer
    
    On Error GoTo ErrorHandler
    
    FullFilePath = FilePath & "\UserDetails.txt"
    
    AryImport = DelimitedTextFileToArray(FullFilePath)
    TotalLines = UBound(AryImport)
    
    ShtMain.WritePersDetails AryImport
    
    ImportPersData = False
Exit Function

ErrorHandler:
    ImportPersData = True
End Function

' ===============================================================
' DelimitedTextFileToArray
' Takes a delimited text file and imports into an array
' ---------------------------------------------------------------
Private Function DelimitedTextFileToArray(FilePath As String) As Variant()
    Dim Delimiter As String
    Dim TextFile As Integer
    Dim FileContent As String
    Dim LineArray() As String
    Dim DataArray() As Variant
    Dim TempArray() As String
    Dim Rows As Integer
    Dim rw As Long, col As Long
    Dim x, y As Integer
    
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

' ===============================================================
' ImportCourseDates
' Imports Course Dates from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ImportCourseDates(FilePath) As Boolean
    Dim AryImport() As Variant
    Dim FullFilePath As String
    Dim TotalLines As Integer
    
    On Error GoTo ErrorHandler
    
    FullFilePath = FilePath & "\CourseDates.txt"
    
    AryImport = DelimitedTextFileToArray(FullFilePath)
    TotalLines = UBound(AryImport)
    
    ShtCourseDates.WriteCourseDates AryImport
    
    ImportCourseDates = False
Exit Function

ErrorHandler:
    ImportCourseDates = True
End Function

' ===============================================================
' ExportTrendData
' Exports Trend data to a text file and saves at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ExportTrendData(FilePath As String) As Boolean
    Dim AryTrendData() As Variant
    Dim CrewCount As Integer
    Dim rw As Integer
    Dim Cl As Integer
    Dim TxtLine As String
    Dim ExpFile As TextStream
    
    On Error GoTo ErrorHandler
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Set ExpFile = FSO.CreateTextFile(FilePath & "\TrendData.txt", True)
    
    AryTrendData = ShtDashboard.GetTrendData
    
    For rw = LBound(AryTrendData) To UBound(AryTrendData)
        TxtLine = ""
        For Cl = 1 To 5
            TxtLine = TxtLine & AryTrendData(rw, Cl) & ";"
        Next
        ExpFile.WriteLine (TxtLine)
    Next
    ExpFile.Close
    
    ExportTrendData = False
    
    Set ExpFile = Nothing
    Set FSO = Nothing
Exit Function

ErrorHandler:
    Set ExpFile = Nothing
    Set FSO = Nothing
    ExportTrendData = True
End Function

' ===============================================================
' ImportTrendData
' Imports Trend Data from a text file stored at the location
' specified.  Returns TRUE if an error occurs.
' ---------------------------------------------------------------
Private Function ImportTrendData(FilePath) As Boolean
    Dim AryImport() As Variant
    Dim FullFilePath As String
    Dim TotalLines As Integer
    
    On Error GoTo ErrorHandler
    
    FullFilePath = FilePath & "\TrendData.txt"
    
    AryImport = DelimitedTextFileToArray(FullFilePath)
    TotalLines = UBound(AryImport)
    
    ShtDashboard.WriteTrendData AryImport
    
    ImportTrendData = False
Exit Function

ErrorHandler:
    ImportTrendData = True
End Function

