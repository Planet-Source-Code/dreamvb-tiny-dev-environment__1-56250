Attribute VB_Name = "ModTools"
Public TMouseButton As Integer
Public tIndex As Integer ' used to hold the index of a command button
Public TKeyCode As Integer
Public dScript As New clsMain
Public TheObjectName As Object

Function Encrypt(lzStr As String) As String
Dim sByte() As Byte
' small simple function to encrypt / decrypt text
    sByte() = StrConv(lzStr, vbFromUnicode)
    
    For I = LBound(sByte) To UBound(sByte)
        sByte(I) = sByte(I) Xor 43
    Next
    
    Encrypt = StrConv(sByte, vbUnicode)
    I = 0
    Erase sByte()
End Function

Public Function IsFileHere(lzFilename As String) As Boolean
    If Dir(lzFilename) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    If Right(FixPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Sub ArangeControl(CtrObj As Object, theForm As Form)
    CtrObj.ZOrder vbBringToFront ' bring the object to the front
    ' code below to center the object on the form designer
    CtrObj.Top = (theForm.ScaleHeight - CtrObj.Height) / 2
    CtrObj.Left = (theForm.ScaleWidth - CtrObj.Width) / 2
    CtrObj.Visible = True ' show the object
    Set TheObjectName = CtrObj ' assign TheObjectName with CtrObj
End Sub

Public Sub tAddControl(Frm As Form, CtrlName As String)
Dim CtrlCount As Integer
Dim mObj As Object
' This function below is uses to add a new control to the form
    CtrlCount = 0
    Select Case UCase(CtrlName)
        Case "T_BUTTON" ' A Command Button
            CtrlCount = Frm.CmdBut.Count
            Load Frm.CmdBut(CtrlCount)
            ArangeControl Frm.CmdBut(CtrlCount), frmWorkArea
        Case "T_IMAGE" ' A Picture Box
            CtrlCount = Frm.PicImg.Count
            Load Frm.PicImg(CtrlCount)
            ArangeControl Frm.PicImg(CtrlCount), frmWorkArea
        Case "T_LABEL" ' A Label
            CtrlCount = Frm.lblA.Count
            Load Frm.lblA(CtrlCount)
            ArangeControl Frm.lblA(CtrlCount), frmWorkArea
        Case "T_TEXT" ' A Text Box
            CtrlCount = Frm.txtA.Count
            Load Frm.txtA(CtrlCount)
            ArangeControl Frm.txtA(CtrlCount), frmWorkArea
        Case "T_LIST"
            CtrlCount = Frm.lstA.Count
            Load Frm.lstA(CtrlCount)
            ArangeControl Frm.lstA(CtrlCount), frmWorkArea
    End Select
End Sub

Public Sub CleanUpAll()
    Set TheObjectName = Nothing
    Erase ProjectData.mCommandButton()
    Erase ProjectData.mlabel()
    Erase ProjectData.mPictureBox()
    Erase ProjectData.mTextBox()
    ProjectData.mSIG = ""
    ProjectData.mVersion = 0
    ProjectData.mFormData.nBackColor = 0
    ProjectData.mFormData.nCaption = ""
    ProjectData.mFormData.nHeight = 0
    ProjectData.mFormData.nStartPosition = 0
    ProjectData.mFormData.nWidth = 0
    TProject.UnitFile = ""
    TProject.ProgLan = ""
    TKeyCode = 0
    Set dScript.DialogObject = Nothing
    dScript.DialogStrName = ""
    dScript.mLanguage = ""
End Sub


