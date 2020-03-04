Attribute VB_Name = "Template_File_Handling"
Option Explicit
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'For Output - When you are opening the text file with this command, you are wanting to create or modify the text file. You will not be able to pull anything from the text file while opening with this mode.
'For Input - When you are opening the text file with this command, you are wanting to extract information from the text file. You will not be able to modify the text file while opening it with this mode.
'For Append - Add new text to the bottom of your text file content.
'FreeFile - Is used to supply a file number that is not already in use. This is similar to referencing Workbook(1) vs. Workbook(2). By using FreeFile, the function will automatically return the next available reference number for your text file.
'Write - This writes a line of text to the file surrounding it with quotations
'Print - This writes a line of text to the file without quotations
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

' Create and fill text file - Option 1
' =====================================
Sub TextFile_Create()
    Dim TextFile As Integer
    Dim FilePath As String
    
    'What is the file path and name for the new text file?
    FilePath = "C:\Users\chris\Desktop\MyFile.txt"
    
    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile
    
    'Open the text file
    Open FilePath For Output As TextFile
    
    'Write some lines of text
    Print #TextFile, "Hello Everyone!"
    Print #TextFile, "I created this file with VBA."
    Print #TextFile, "Goodbye"
    
    'Save & Close Text File
    Close TextFile
    
End Sub

' Create and fill text file - Option 2
' =====================================
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

' Read all content from Text File
' =====================================
Sub TextFile_PullData()
    Dim TextFile As Integer
    Dim FilePath As String
    Dim FileContent As String
    
    'File Path of Text File
    FilePath = "C:\Users\chris\Desktop\MyFile.txt"
    
    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile
    
    'Open the text file
    Open FilePath For Input As TextFile
    
    'Store file content inside a variable
    FileContent = Input(LOF(TextFile), TextFile)
    
    'Report Out Text File Contents
    MsgBox FileContent
    
    'Close Text File
    Close TextFile
    
End Sub

'Read line by line from text file - simple
'========================================
Sub ReadFileLineByLine()
    Dim my_file As Integer
    Dim text_line As String
    Dim file_name As String
    Dim i As Integer

    file_name = "C:\text_file.txt"

    my_file = FreeFile()
    Open file_name For Input As my_file

    i = 1

    While Not EOF(my_file)
        Line Input #my_file, text_line
        Cells(i, "A").Value = text_line
        i = i + 1
    Wend
End Sub

'Read content line by line from text file into an array
'======================================================
Function DelimitedTextFileToArray(FilePath As String) As String()
    Dim Delimiter As String
    Dim TextFile As Integer
    Dim FileContent As String
    Dim LineArray() As String
    Dim DataArray() As String
    Dim TempArray() As String
    Dim rw As Long, col As Long
    Dim x As Integer
    Dim y As Integer
    
    'Inputs
    Delimiter = ";"
    rw = 0
    
    'Open the text file in a Read State
    TextFile = FreeFile
    Open FilePath For Input As TextFile
    
    'Store file content inside a variable
    FileContent = Input(LOF(TextFile), TextFile)
    
    'Close Text File
    Close TextFile
    
    'Separate Out lines of data
    LineArray() = Split(FileContent, vbCrLf)
    
    'Read Data into an Array Variable
    For x = LBound(LineArray) To UBound(LineArray)
        If Len(Trim(LineArray(x))) <> 0 Then
            'Split up line of text by delimiter
            TempArray = Split(LineArray(x), Delimiter)
            
            'Determine how many columns are needed
            col = UBound(TempArray)
            
            'Re-Adjust Array boundaries
            ReDim Preserve DataArray(col, rw)
            
            'Load line of data into Array variable
            For y = LBound(TempArray) To UBound(TempArray)
                DataArray(y, rw) = TempArray(y)
            Next y
        End If
    
        'Next line
        rw = rw + 1
    
    Next x
    DelimitedTextFileToArray = DataArray()
    
End Function
