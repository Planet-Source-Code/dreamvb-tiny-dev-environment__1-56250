VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub PutPixelA(ByVal tHDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    SetPixel tHDC, X, Y, Color
End Sub

Public Sub DrawLine(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Colour As Long)
    Form1.Line (x1, y1)-(x2, y2), Colour, B
End Sub

Public Sub Printf(S As String)
    frmWorkArea.ForeColor = frmWorkArea.ForeColorf
    frmWorkArea.Print S
End Sub

Function TextOutA(ByVal tHDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Text As String) As Long
    TextOutA = TextOut(tHDC, X, Y, Text, Len(Text))
End Function

Function GetHDC(ByVal Hangle As Long) As Long
    GetHDC = GetDC(Hangle)
End Function

Function GetPixelA(ByVal tHDC As Long, ByVal X As Long, ByVal Y As Long)
    GetPixelA = GetPixel(tHDC, X, Y)
End Function
