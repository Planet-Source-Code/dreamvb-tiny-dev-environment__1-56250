Attribute VB_Name = "ModAPI"
' Window API calls
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

'Windows API calls
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function IsPwrShutdownAllowed Lib "Powrprof.dll" () As Long
Public Declare Function IsPwrSuspendAllowed Lib "Powrprof.dll" () As Long
Public Declare Function IsPwrHibernateAllowed Lib "Powrprof.dll" () As Long
' Drive and file API calls
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' Internet / Network API calls
Public Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function NetUserChangePassword Lib "NETAPI32.DLL" (ByVal domainname As String, ByVal userName As String, ByVal oldpassword As String, ByVal newpassword As String) As Long
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
'Graphic API Calls
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
' window stayle consts
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
' Browse folder consts
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
' windows version consts
Private Const VER_PLATFORM_WIN32_NT = 2

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum TSpecialFolders
    DM_DESKTOP = &H0
    DM_PROGRAMS = &H2
    DM_Controls = &H3
    DM_PRINTERS = &H4
    DM_PERSONAL = &H5
    DM_FAVORITES = &H6
    DM_STARTUP = &H7
    DM_RECENT = &H8
    DM_SENDTO = &H9
    DM_BITBUCKET = &HA
    DM_STARTMENU = &HB
    DM_DESKTOPDIRECTORY = &H10
    DM_DRIVES = &H11
    DM_NETWORK = &H12
    DM_NETHOOD = &H13
    DM_FONTS = &H14
    DM_TEMPLATES = &H15
End Enum

Public Osver As OSVERSIONINFO
Public nPos As POINTAPI

Public Function GetShortPath(lzPathName As String) As String
    Dim iRet As Long, StrA As String
    StrA = String$(512, vbNullChar)
    iRet = GetShortPathName(lzPathName, StrA, 164)
    GetShortPath = Left$(StrA, iRet)
    StrA = ""
    iRet = 0
End Function

Public Function isWinNT() As Boolean
    Osver.dwOSVersionInfoSize = Len(Osver)
    GetVersionEx Osver
    isWinNT = (Osver.dwPlatformId) = VER_PLATFORM_WIN32_NT
End Function

Function GetFolder(ByVal hwndOwner As Long, mTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim OffSet As Integer
    If Len(mTitle) = 0 Then mTitle = "Look in:"
    bInf.hOwner = hwndOwner
    bInf.lpszTitle = mTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        OffSet = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, OffSet - 1)
    End If
End Function

Public Function tFlashWindow(mWnd As Long, mIntVal As Long)
   FlashWindow mWnd, mIntVal
End Function

Public Function GetUser() As String
Dim S As String, iRet As Long
    S = Space(165)
    iRet = GetUserName(S, 165)
    If iRet <> 1 Then
        GetUser = ""
        S = ""
        Exit Function
    Else
        GetUser = Left(S, InStr(S, Chr(0)) - 1)
        S = ""
    End If
End Function

Public Function SysComputerName() As String
Dim S As String, iRet As Long
    S = Space(160)
    iRet = GetComputerName(S, 160)
    
    If iRet <> 1 Then
        SysComputerName = ""
        S = ""
        Exit Function
    Else
        SysComputerName = Left(S, InStr(S, Chr(0)) - 1)
        S = ""
    End If
    
End Function

Public Function DMGetSystemPath() As String
Dim Strbuff As String
    Strbuff = String(255, Chr(0))
    GetSystemDirectory Strbuff, 255
    DMGetSystemPath = Left(Strbuff, InStr(Strbuff, Chr(0)) - 1)
    Strbuff = ""
End Function

Public Function DMGetTempPath() As String
Dim Strbuff As String
    Strbuff = String(255, Chr(0))
    GetTempPath 255, Strbuff
    DMGetTempPath = Left(Strbuff, InStr(1, Strbuff, Chr(0)) - 1)
    Strbuff = ""
End Function

Public Function DMGetWindowsPath() As String
Dim Strbuff As String
    Strbuff = String(255, Chr(0))
    GetWindowsDirectory Strbuff, 255
    DMGetWindowsPath = Left(Strbuff, InStr(1, Strbuff, Chr(0)) - 1)
    Strbuff = ""
End Function
