VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public HasError As Boolean
Public DialogStrName As String
Public DialogObject As Form

Private TLanguage As String
Public sControl As ScriptControl
Private Graphics As clsGraphics
Private Windows As clswindows
Private FileSys As clsFiles
Private TDrive As clsDrive
Private TNetwork As clsNetwork
Private TTools As clsTools
Private TIniFile As clsINI

Public Sub SetupControl()
    Set sControl = New ScriptControl
    Set Graphics = New clsGraphics
    Set Windows = New clswindows
    Set FileSys = New clsFiles
    Set TDrive = New clsDrive
    Set TNetwork = New clsNetwork
    Set TTools = New clsTools
    Set TIniFile = New clsINI
    
    sControl.Language = TLanguage
    sControl.AddObject DialogStrName, DialogObject, True
    'sControl.AddObject "Window", frmWindow, True
    sControl.AddObject "Graphics", Graphics, True
    sControl.AddObject "Windows", Windows, True
    sControl.AddObject "FileSys", FileSys, True
    sControl.AddObject "TDrive", TDrive, True
    sControl.AddObject "TNetwork", TNetwork, True
    sControl.AddObject "TTools", TTools, True
    sControl.AddObject "TIniFile", TIniFile, True
    
    'sControl.AddObject "bTimer", frmWindow.bTimer, True
    sControl.AllowUI = True
    
    TTools.Hangle = DialogObject.hwnd
End Sub

Public Sub mRunProc(StrCode As String)
    On Error Resume Next
    sControl.Run StrCode
End Sub

Public Sub Reset()
On Error Resume Next
    sControl.Reset
End Sub
Public Sub RunCode(StrCode As String)
On Error GoTo CodeErr
    If Len(StrCode) = 0 Then Exit Sub
    
    sControl.AddCode StrCode
    Exit Sub
CodeErr:
    If Err Then
    MsgBox Err.Description
    
        HasError = True
    End If
    
End Sub

Public Property Get mLanguage() As String
    mLanguage = TLanguage
End Property

Public Property Let mLanguage(ByVal vNewValue As String)
    TLanguage = vNewValue
End Property
