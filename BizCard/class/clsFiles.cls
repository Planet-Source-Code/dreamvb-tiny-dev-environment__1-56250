VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PakAddFile() As String

Function AppendFile(Filename As String, Text As String)
Dim iFile As Long
    iFile = FreeFile
    Open Filename For Append As #iFile
        Print #iFile, Text
    Close #iFile
End Function

Public Function WriteFile(Filename As String, DataBuff As String)
Dim iFile As Long
    iFile = FreeFile
    Open Filename For Binary As #iFile
        Put #iFile, , DataBuff
    Close #iFile
    
End Function
Public Function OpenFile(Filename As String) As String
Dim iFile As Long, StrB As String
    iFile = FreeFile
    Open Filename For Binary As #iFile
        StrB = Space(LOF(iFile))
        Get #iFile, , StrB
    Close #iFile
    
    OpenFile = StrB
    StrB = ""
    
End Function

Function CreateFolder(lzPath As String)
Dim NewDir As String, NewPath As String, mDir As Variant, I As Long
On Error Resume Next
    If Trim(Len(lzPath)) <= 0 Then Exit Function
    If Right(lzPath, 1) = "\" Then lzPath = Left(lzPath, Len(lzPath) - 1)
    
    mDir = Split(lzPath, "\")
    
    For I = LBound(mDir) To UBound(mDir)
        NewDir = mDir(I)
        If Not Len(NewDir) > 0 Then
            Exit Function
        Else
            NewPath = NewPath & mDir(I) & "\"
        End If
        
        If Dir(NewPath, vbDirectory) = "" Then
            MkDir NewPath
        End If
    Next
    
    I = 0
    NewPath = ""
    NewDir = ""
    Erase mDir
End Function

Function DeleteFolder(lzPath As String) As Long
On Error Resume Next
If FolderExists(lzPath) = False Then DeleteFolder = 0: Exit Function
    RmDir lzPath
    DeleteFolder = 1
End Function

Function GetFileAttributes(Filename As String) As Long
On Error Resume Next
    GetFileAttributes = GetAttr(Filename)
    If Err Then GetFileAttributes = -1
End Function

Function GetFileSize(Filename As String) As Long
    If FindFile(Filename) = 0 Then GetFileSize = 0: Exit Function
    GetFileSize = FileLen(Filename)
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function FolderExists(lzPath As String) As Integer
    If Dir(lzPath, vbDirectory) = "" Then FolderExists = 0 Else FolderExists = 1
End Function

Function FindFile(lzFilename As String) As Integer
    If Len(Trim(lzFilename)) = 0 Then Exit Function
    If Dir(lzFilename) = "" Then FindFile = 0: Exit Function Else FindFile = 1
End Function

Function FileDateTimeA(Filename As String) As Date
    FileDateTimeA = FileDateTime(Filename)
End Function

Function DeleteFile(Filename As String) As Long
On Error GoTo DelErr
    Kill Filename
    DeleteFile = 1
    Exit Function
DelErr:
    If Err Then DeleteFile = 0
End Function

Function CopyFile(File1 As String, File2 As String) As Long
On Error GoTo CopyErr
    FileCopy File1, File2
    CopyFile = 1
    Exit Function
CopyErr:
    If Err Then CopyFile = 0
End Function

Function RunControlPanelApp(ByVal ProgName As String) As Long
    If Len(Trim(ProgName)) = 0 Then Exit Function
    RunControlPanelApp = WinExec("Control.exe " & ProgName, 1)
End Function
Function Run(Filename As String) As Long
    Run = WinExec(Filename, 1)
End Function

Function MoveFileA(ByVal File1 As String, ByVal File2 As String) As Long
    MoveFileA = MoveFile(File1, File2)
End Function
