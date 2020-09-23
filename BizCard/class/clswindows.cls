VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clswindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function BeepA()
    Beep
End Function
'mciSendStringA "Set CDAudio Door Open Wait", "", 0, 0
'mciSendStringA "Set CDAudio Door Closed Wait","",0,0

Function MciSendStringA(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long)
    MciSendStringA = mciSendString(lpstrCommand, lpstrReturnString, uReturnLength, hwndCallback)
End Function
Function ActivateWindow(ByVal lzWindowName As String) As Long
Dim WndHandle As Long
    WndHandle = FindWindow(vbNullString, lzWindowName)
    SetForegroundWindow WndHandle
End Function

Function AppPath() As String
Dim lzPath As String
    lzPath = App.Path
    If Right(lzPath, 1) = "\" Then lzPath = App.Path Else lzPath = App.Path & "\"
    AppPath = lzPath
    lzPath = ""
End Function

Function AlphaBlend(Hangle As Long, Optional AlphaBlendValue As Integer = 165)
Dim wnd As Long
    wnd = GetWindowLong(Hangle, GWL_EXSTYLE)
    wnd = wnd Or WS_EX_LAYERED
    SetWindowLong Hangle, GWL_EXSTYLE, wnd
    SetLayeredWindowAttributes Hangle, 0, AlphaBlendValue, &H2
    wnd = 0
End Function

Public Function SetWindowFocus(ByVal Hangle As Long)
    SetForegroundWindow Hangle
End Function

Public Function MessageBox(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title = "Message Box") As Long
On Error Resume Next
    MessageBox = MsgBox(Prompt, Buttons, Title)
End Function

Public Function FlashWindow(hwnd As Long, mInterval As Long)
    tFlashWindow hwnd, mInterval
End Function

Public Function FindWindowA(WndClsName As String, WndName As String)
    If Len(WndClsName) = 0 Then WndClsName = vbNullString
    FindWindowA = FindWindow(WndClsName, WndName)
End Function

Public Function CloseWindowA(WndHangle As Long)
    CloseWindow WndHangle
End Function

Public Function BrowseForFolder(Hangle As Long, Optional Title As String)
   BrowseForFolder = GetFolder(Hangle, Title)
End Function
Function GetActiveWindowA() As Long
    GetActiveWindowA = GetActiveWindow
End Function
Function GetForegroundWindowA() As Long
    GetForegroundWindowA = GetForegroundWindow()
End Function

Function GetComputerNameA() As String
    GetComputerNameA = SysComputerName
End Function

Function GetUserNameA() As String
    GetUserNameA = GetUser
End Function

Function GetMousePos()
    GetCursorPos nPos
End Function

Public Sub SetMousePos(ByVal X As Long, Y As Long)
    SetCursorPos X, Y
End Sub

Function GetTickCountA() As Long
    GetTickCountA = GetTickCount&
End Function

Function GetOSVerType() As Integer
    Osver.dwOSVersionInfoSize = Len(Osver)
    GetVersionEx Osver
    GetOSVerType = Osver.dwPlatformId
End Function

Public Function GetSpecialFolderLocation(ByVal bsSpecialFolder As String) As String
Dim TFolder As TSpecialFolders
    Dim spFolder As Long
    Dim Strbuff As String
    Dim RetVal As Long
    Dim IDL As ITEMIDLIST
    
    Select Case UCase(bsSpecialFolder)
        Case "DESKTOP"
            spFolder = &H0
        Case "PROGRAMS"
            spFolder = &H2
        Case "CONTROLS"
            spFolder = &H3
        Case "PRINTERS"
            spFolder = &H4
        Case "PERSONAL"
            spFolder = &H5
        Case "FAVORITES"
            spFolder = &H6
        Case "STARTUP"
            spFolder = &H7
        Case "RECENT"
            spFolder = &H8
        Case "SENDTO"
            spFolder = &H9
        Case "BITBUCKET"
            spFolder = &HA
        Case "STARTMENU"
            spFolder = &HB
        Case "DESKTOPDIRECTORY"
            spFolder = &H10
        Case "DRIVES"
            spFolder = &H11
        Case "NETWORK"
            spFolder = &H12
        Case "NETHOOD"
            spFolder = &H13
        Case "FONTS"
            spFolder = &H14
        Case "TEMPLATES"
            spFolder = &H15
        Case "SYSTEM"
            GetSpecialFolderLocation = DMGetSystemPath
            Exit Function
        Case "TEMP"
            GetSpecialFolderLocation = DMGetTempPath
            Exit Function
        Case "WINDOWS"
           GetSpecialFolderLocation = DMGetWindowsPath
           Exit Function
        Case Else
            GetSpecialFolderLocation = ""
            Exit Function
    End Select
    
    RetVal = SHGetSpecialFolderLocation(100, spFolder, IDL)
    
    If RetVal = 0 Then
        Strbuff = String(512, Chr(0))
        RetVal = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Strbuff)
        GetSpecialFolderLocation = Left(Strbuff, InStr(Strbuff, Chr(0)) - 1)
    End If
    
    bsSpecialFolder = ""
    RetVal = 0
    Strbuff = ""
    
End Function

Function isShutdownAllowed() As Long
    isShutdownAllowed = IsPwrShutdownAllowed
End Function

Function isSuspendAllowed() As Long
    isSuspendAllowed = IsPwrSuspendAllowed
End Function

Function isHibernateAllowed() As Long
    isHibernateAllowed = IsPwrHibernateAllowed
End Function

Function RunDialog(Hangle As Long, Optional ByVal Title As String = "Run", Optional ByVal Prompt As String = "Enter the name of the program to run")
    If isWinNT Then
         SHRunDialog Hangle, 0, 0, StrConv(Title, vbUnicode), StrConv(Prompt, vbUnicode), 0
    Else
         SHRunDialog Hangle, 0, 0, Title, Prompt, 0
    End If
End Function

Function Pause(ByVal Millisecond As Long)
    Sleep Millisecond
End Function
Function GetWindowPosition(ByVal Hangle As Long) As Variant()
Dim Holder(3) As Variant

Dim mRect As RECT
    GetWindowRect Hangle, mRect
    Holder(0) = mRect.Left ' x pos
    Holder(1) = mRect.Top ' pos
    Holder(2) = mRect.Right ' width
    Holder(3) = mRect.Bottom ' height
    GetWindowPosition = Holder
    
End Function
Function SetWindowPosition(ByVal Hangle As Long, ByVal X, ByVal Y, ByVal mHeight As Long, ByVal mWidth As Long) As Long
    SetWindowPosition = MoveWindow(Hangle, X, Y, mWidth, mHeight, 1)
End Function
Function SetWindowTextA(ByVal Hangle As Long, ByVal lText As String)
    SetWindowTextA = SetWindowText(Hangle, lText)
End Function
Function SendKeysA(TKeys As String, Optional ByVal Wait As Integer)
    SendKeys TKeys, Wait
End Function
Public Property Get MouseX() As Long
    MouseX = nPos.X
End Property

Public Property Get MouseY() As Long
    MouseY = nPos.Y
End Property

Public Property Get MouseButton() As Integer
    MouseButton = TMouseButton
End Sub

Public Property Get Index() As Integer
    Index = tIndex
End Property
