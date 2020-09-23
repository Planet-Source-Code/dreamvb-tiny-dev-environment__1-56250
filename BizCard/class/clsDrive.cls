VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function DriveType(ByVal mDriveLetter As String) As Long
    If Len(Trim(mDriveLetter)) = 0 Then: DriveType = 1: Exit Function
    DriveType = GetDriveType(mDriveLetter)
End Function
