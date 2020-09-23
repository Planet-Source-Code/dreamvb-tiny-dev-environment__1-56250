VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function MapNetworkDriveDlg(ByVal hangle As Long) As Long
   WNetConnectionDialog hangle, 1
End Function

Function DisconnectNetworkDriveDlg(ByVal hangle As Long)
    WNetDisconnectDialog hangle, 1
End Function

Public Function DownloadFile(ByVal URL As String, ByVal LocalFilename As String) As Boolean
    Dim RetVal As Long
    RetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If RetVal = 0 Then DownloadFile = True
End Function

Function isAdmin() As Long
    isAdmin = IsNTAdmin(ByVal 0&, ByVal 0&)
End Function
