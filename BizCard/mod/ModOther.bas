Attribute VB_Name = "ModOther"
Enum ViewOp
    iCodeView = 0
    iDialogView = 1
End Enum

Public TheObjectName As Object
Public inIde As Boolean
Public ObjectSelected As ViewOp
Public Modified As Boolean
Public ButtonPressed As Integer
Public OldWindowState As Integer
Public FirstTimeLoad As Boolean
Public dScript As New clsMain
Public ProjectFileToOpen As String

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

Function RemoveComments(StrString As String) As String
Dim CodeLine As String, tCode As String
Dim vStr As Variant
Dim ipos As Long, I As Long

    vStr = Split(StrString, vbCrLf)
    For I = LBound(vStr) To UBound(vStr)
        tCode = vStr(I)
        ipos = InStr(1, vStr(I), "'", vbTextCompare)
        
        If ipos > 0 Then
            vStr(I) = Mid(tCode, 1, ipos - 1)
        End If
    
        CodeLine = CodeLine & Trim(vStr(I)) & vbCrLf
    Next
    CodeLine = Left(CodeLine, Len(CodeLine) - 2)
    ipos = 0
    I = 0
    tCode = ""
    Erase vStr
    RemoveComments = CodeLine
    CodeLine = ""
End Function

Private Sub ArangeControl(CtrObj As Object, theForm As Form)
    CtrObj.ZOrder vbBringToFront ' bring the object to the front
    ' code below to center the object on the form designer
    CtrObj.Top = (theForm.ScaleHeight - CtrObj.Height) / 2
    CtrObj.Left = (theForm.ScaleWidth - CtrObj.Width) / 2
    CtrObj.Visible = True ' show the object
    Set TheObjectName = CtrObj ' assign TheObjectName with CtrObj
    frmWorkArea.MakeSelection True ' show the objects selection
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

Public Sub DialogRun(Frm As Form)
    inIde = False
    frmWorkArea.HideSelection
    frmWorkArea.Cls
    Set frmWorkArea.Picture = Nothing
    DoEvents
    
End Sub

Public Sub CleanUpAll()
    Set TheObjectName = Nothing
    Modified = False
    ButtonPressed = 0
    
    Erase ProjectData.mCommandButton()
    Erase ProjectData.mlabel()
    Erase ProjectData.mPictureBox()
    Erase ProjectData.mTextBox()
    Erase ProjectData.nListBox()
    ProjectData.mSIG = ""
    ProjectData.mVersion = 0
    ProjectData.mFormData.nBackColor = 0
    ProjectData.mFormData.nCaption = ""
    ProjectData.mFormData.nHeight = 0
    ProjectData.mFormData.nStartPosition = 0
    ProjectData.mFormData.nWidth = 0
    ProjectFolder = ""
    ProjectName = ""
    TProject.FormFile = ""
    TProject.ProjectTitle = ""
    TProject.UnitFile = ""
    frmCode.txtCode.Text = ""
End Sub
