VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public TStack As New Collection
Public hangle As Long
Const mText = 2

' Class used for string and math function, subs
Function CountChr(ByVal Range As String, ByVal Criteria As String) As Long
Dim I As Integer
Dim Cnt As Long

If Len(Trim(Range)) = 0 Then CountIF = 0: Exit Function

    Do While I < Len(Range)
        I = I + 1
        If Mid(Range, I, 1) = Criteria Then
            Cnt = Cnt + 1
        End If
        DoEvents
    Loop
    
    CountChr = Cnt
    Cnt = 0
    Range = ""
    Criteria = ""
    
End Function

Function StrToVal(StrString As String) As Integer
    ' Converts a string to a integer
    StrToVal = CInt(StrString)
End Function

Function StrFormatDateTime(lExpression As String, Optional ByVal lFormat As Variant) As Variant
    StrFormatDateTime = Format(lExpression, lFormat)
End Function

Function TColor(ByVal lColor As Integer) As Long
    If lColor < 0 Then
        TColor = 0
        Exit Function
    ElseIf lColor >= 15 Then
        TColor = vbWhite
        Exit Function
    Else
        TColor = QBColor(lColor)
    End If
End Function

Function GetEnvVar(sName As String) As String
'This is used to return the string assigned to an environment variable
Dim iRet As Long, sBuff As String
    sBuff = Space(255)
    iRet = GetEnvironmentVariable(sName, sBuff, 255)
    
    If Not iRet <> 0 Then
        GetEnvVar = ""
        sBuff = ""
        Exit Function
    Else
        GetEnvVar = Left(sBuff, iRet)
        iRet = 0
        sBuff = ""
    End If
End Function

Function SetEnvVar(ByVal sName As String, ByVal sValue As String) As Long
    'This is used to set an environment variable
    SetEnvVar = SetEnvironmentVariable(sName, sValue)
End Function

Public Sub Swap(a, b)
Dim Temp
    Temp = b
    b = a
    a = Temp
    Temp = 0
End Sub

Function Power(ByVal iNum As Variant, ByVal iCount As Variant)
    Power = (iNum * iCount)
End Function

Function Prompt(Optional ByVal aPrompt As String, Optional ByVal Title As String) As String
    Prompt = InputBox(aPrompt, Title)
End Function
'end

Function GetClip(Optional ByVal zFormatType As Integer = 1) As String
    If (zFormatType < 1) Or (zFormatType > 2) Then zFormatType = 1

    If zFormatType <= 1 Then
        GetClip = Clipboard.GetText(vbCFText)
    End If
    
    If zFormatType = 2 Then
        GetClip = Clipboard.GetText(vbCFRTF)
    End If
    
End Function

Function SetClip(ByVal Strbuff As String, Optional ByVal zFormatType As Integer = 1) As Integer

    If (zFormatType < 1) Or (zFormatType > 2) Then zFormatType = 1
    Clipboard.Clear
    If zFormatType <= 1 Then
        Clipboard.SetText Strbuff, vbCFText
        SetClip = 1
    End If
    
    If zFormatType = 2 Then
        Clipboard.SetText Strbuff, vbCFRTF
        SetClip = 1
    End If
    
End Function

Public Sub SaveSettingA(ByVal tAppName As String, ByVal tSelection As String, ByVal tKey As String, ByVal tSetting As String)
    SaveSetting tAppName, tSelection, tKey, tSetting
End Sub

Public Function GetSettingA(ByVal tAppName As String, ByVal tSelection As String, ByVal tKey As String, ByVal tDefault As String) As String
    GetSettingA = GetSetting(tAppName, tSelection, tKey, tDefault)
End Function

Public Property Get bsText() As Integer
    bsText = 1
End Property

Public Property Get bsRTF() As Integer
    bsRTF = 2
End Property

