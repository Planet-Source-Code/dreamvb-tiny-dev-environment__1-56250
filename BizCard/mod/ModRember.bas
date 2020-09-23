Attribute VB_Name = "ModRember"
Private Type TForm
    nCaption As String
    nHeight As Long
    nWidth As Long
    nBackColor As Long
    nStartPosition As Integer
End Type

Private Type TCommandButtonA
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nCaption As String
    nBackColor As Long
    nFontName As String
    nFontSize As Long
    nFontBold As Boolean
    nFontItalic As Boolean
    nFontUnderline As Boolean
End Type

Private Type TPictureA
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nBackColor As Long
    nBorderStyle As Integer
    nAutoSize As Boolean
End Type

Private Type TListBox
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nBackColor As Long
    nForeColor As Long
    nFontName As String
    nFontSize As Long
    nFontBold As Boolean
    nFontItalic As Boolean
    nFontUnderline As Boolean
End Type

Private Type TLabelA
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nCaption As String
    nBackColor As Long
    nForeColor As Long
    nAlignment As Integer
    nAutoSize As Integer
    nFontName As String
    nFontSize As Long
    nFontBold As Boolean
    nFontItalic As Boolean
    nFontUnderline As Boolean
End Type

Private Type TTextBoxA
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nBackColor As Long
    nForeColor As Long
    nFontName As String
    nFontSize As Long
    nFontBold As Boolean
    nFontItalic As Boolean
    nFontUnderline As Boolean
    nText As String
End Type

Private Type TRember
    mFormData As TForm
    mCommandButton() As TCommandButtonA
    mPictureBox() As TPictureA
    mlabel() As TLabelA
    mTextBox() As TTextBoxA
    mListBox() As TListBox
End Type

Public TBackup As TRember

Public Sub RemberFormData(Frm As Form)
Dim CtrlCounter(4) As Integer, Counter As Long, FrmCtrName As String, ObjRember As Object
Dim nFile As Long


    ReDim Preserve TBackup.mCommandButton(0)
    ReDim Preserve TBackup.mlabel(0)
    ReDim Preserve TBackup.mPictureBox(0)
    ReDim Preserve TBackup.mTextBox(0)
    ReDim Preserve TBackup.mListBox(0)
    
    TBackup.mFormData.nBackColor = Frm.BackColor
    TBackup.mFormData.nCaption = Frm.Caption
    TBackup.mFormData.nWidth = Frm.Width
    TBackup.mFormData.nHeight = Frm.Height
    TBackup.mFormData.nStartPosition = 1
    
    For Counter = 7 To Frm.Controls.Count - 1
        FrmCtrName = TypeName(Frm.Controls(Counter))
        Set ObjRember = Frm.Controls(Counter)
        Select Case UCase(FrmCtrName)
            Case "COMMANDBUTTON"
                CtrlCounter(0) = CtrlCounter(0) + 1
                ReDim Preserve TBackup.mCommandButton(CtrlCounter(0))
                TBackup.mCommandButton(CtrlCounter(0)).nWidth = ObjRember.Width
                TBackup.mCommandButton(CtrlCounter(0)).nHeight = ObjRember.Height
                TBackup.mCommandButton(CtrlCounter(0)).nLeft = ObjRember.Left
                TBackup.mCommandButton(CtrlCounter(0)).nTop = ObjRember.Top
                TBackup.mCommandButton(CtrlCounter(0)).nBackColor = ObjRember.BackColor
                TBackup.mCommandButton(CtrlCounter(0)).nCaption = ObjRember.Caption
                TBackup.mCommandButton(CtrlCounter(0)).nFontBold = ObjRember.FontBold
                TBackup.mCommandButton(CtrlCounter(0)).nFontItalic = ObjRember.FontItalic
                TBackup.mCommandButton(CtrlCounter(0)).nFontUnderline = ObjRember.FontUnderline
                TBackup.mCommandButton(CtrlCounter(0)).nFontName = ObjRember.FontName
                TBackup.mCommandButton(CtrlCounter(0)).nFontSize = ObjRember.FontSize
            Case "PICTUREBOX"
                CtrlCounter(1) = CtrlCounter(1) + 1
                ReDim Preserve TBackup.mPictureBox(CtrlCounter(1))
                TBackup.mPictureBox(CtrlCounter(1)).nAutoSize = ObjRember.AutoSize
                TBackup.mPictureBox(CtrlCounter(1)).nBackColor = ObjRember.BackColor
                TBackup.mPictureBox(CtrlCounter(1)).nBorderStyle = ObjRember.BorderStyle
                TBackup.mPictureBox(CtrlCounter(1)).nWidth = ObjRember.Width
                TBackup.mPictureBox(CtrlCounter(1)).nHeight = ObjRember.Height
                TBackup.mPictureBox(CtrlCounter(1)).nTop = ObjRember.Top
                TBackup.mPictureBox(CtrlCounter(1)).nLeft = ObjRember.Left
            Case "LABEL"
                CtrlCounter(2) = CtrlCounter(2) + 1
                ReDim Preserve TBackup.mlabel(CtrlCounter(2))
                TBackup.mlabel(CtrlCounter(2)).nAlignment = ObjRember.Alignment
                TBackup.mlabel(CtrlCounter(2)).nAutoSize = ObjRember.AutoSize
                TBackup.mlabel(CtrlCounter(2)).nBackColor = ObjRember.BackColor
                TBackup.mlabel(CtrlCounter(2)).nCaption = ObjRember.Caption
                TBackup.mlabel(CtrlCounter(2)).nFontBold = ObjRember.FontBold
                TBackup.mlabel(CtrlCounter(2)).nFontItalic = ObjRember.FontItalic
                TBackup.mlabel(CtrlCounter(2)).nFontUnderline = ObjRember.FontUnderline
                TBackup.mlabel(CtrlCounter(2)).nFontName = ObjRember.FontName
                TBackup.mlabel(CtrlCounter(2)).nFontSize = ObjRember.FontSize
                TBackup.mlabel(CtrlCounter(2)).nForeColor = ObjRember.ForeColor
                TBackup.mlabel(CtrlCounter(2)).nWidth = ObjRember.Width
                TBackup.mlabel(CtrlCounter(2)).nHeight = ObjRember.Height
                TBackup.mlabel(CtrlCounter(2)).nTop = ObjRember.Top
                TBackup.mlabel(CtrlCounter(2)).nLeft = ObjRember.Left
            Case "TEXTBOX"
                CtrlCounter(3) = CtrlCounter(3) + 1
                ReDim Preserve TBackup.mTextBox(CtrlCounter(3))
                TBackup.mTextBox(CtrlCounter(3)).nForeColor = ObjRember.ForeColor
                TBackup.mTextBox(CtrlCounter(3)).nBackColor = ObjRember.BackColor
                TBackup.mTextBox(CtrlCounter(3)).nFontBold = ObjRember.FontBold
                TBackup.mTextBox(CtrlCounter(3)).nFontItalic = ObjRember.FontItalic
                TBackup.mTextBox(CtrlCounter(3)).nFontUnderline = ObjRember.FontUnderline
                TBackup.mTextBox(CtrlCounter(3)).nFontName = ObjRember.FontName
                TBackup.mTextBox(CtrlCounter(3)).nFontSize = ObjRember.FontSize
                TBackup.mTextBox(CtrlCounter(3)).nWidth = ObjRember.Width
                TBackup.mTextBox(CtrlCounter(3)).nHeight = ObjRember.Height
                TBackup.mTextBox(CtrlCounter(3)).nTop = ObjRember.Top
                TBackup.mTextBox(CtrlCounter(3)).nLeft = ObjRember.Left
                TBackup.mTextBox(CtrlCounter(3)).nText = ObjRember.Text
            Case "LISTBOX"
                CtrlCounter(4) = CtrlCounter(4) + 1
                ReDim Preserve TBackup.mListBox(CtrlCounter(4))
                TBackup.mListBox(CtrlCounter(4)).nForeColor = ObjRember.ForeColor
                TBackup.mListBox(CtrlCounter(4)).nBackColor = ObjRember.BackColor
                TBackup.mListBox(CtrlCounter(4)).nFontBold = ObjRember.FontBold
                TBackup.mListBox(CtrlCounter(4)).nFontItalic = ObjRember.FontItalic
                TBackup.mListBox(CtrlCounter(4)).nFontUnderline = ObjRember.FontUnderline
                TBackup.mListBox(CtrlCounter(4)).nFontName = ObjRember.FontName
                TBackup.mListBox(CtrlCounter(4)).nFontSize = ObjRember.FontSize
                TBackup.mListBox(CtrlCounter(4)).nWidth = ObjRember.Width
                TBackup.mListBox(CtrlCounter(4)).nHeight = ObjRember.Height
                TBackup.mListBox(CtrlCounter(4)).nTop = ObjRember.Top
                TBackup.mListBox(CtrlCounter(4)).nLeft = ObjRember.Left
                
        End Select
    Next

    Set ObjRember = Nothing
    Erase CtrlCounter()
    Counter = 0
    FrmCtrName = ""
    
End Sub


Public Sub RestoreData(Frm As Form)
On Error Resume Next
Dim I As Integer
    
    Frm.Caption = TBackup.mFormData.nCaption
    Frm.BackColor = TBackup.mFormData.nBackColor
    Frm.Width = TBackup.mFormData.nWidth
    Frm.Height = TBackup.mFormData.nHeight


    For I = 1 To UBound(TBackup.mCommandButton)
        Set ObjRember = Frm.CmdBut(I)
        ObjRember.Top = TBackup.mCommandButton(I).nTop
        ObjRember.Left = TBackup.mCommandButton(I).nLeft
        ObjRember.Width = TBackup.mCommandButton(I).nWidth
        ObjRember.Height = TBackup.mCommandButton(I).nHeight
        ObjRember.BackColor = TBackup.mCommandButton(I).nBackColor
        ObjRember.Caption = TBackup.mCommandButton(I).nCaption
        ObjRember.FontBold = TBackup.mCommandButton(I).nFontBold
        ObjRember.FontItalic = TBackup.mCommandButton(I).nFontItalic
        ObjRember.FontUnderline = TBackup.mCommandButton(I).nFontUnderline
        ObjRember.FontName = TBackup.mCommandButton(I).nFontName
        ObjRember.FontSize = TBackup.mCommandButton(I).nFontSize
    Next
    I = 0
    
    For I = 1 To UBound(TBackup.mlabel)
        Set ObjRember = Frm.lblA(I)
        ObjRember.Top = TBackup.mlabel(I).nTop
        ObjRember.Left = TBackup.mlabel(I).nLeft
        ObjRember.Width = TBackup.mlabel(I).nWidth
        ObjRember.Height = TBackup.mlabel(I).nHeight
        ObjRember.ForeColor = TBackup.mlabel(I).nForeColor
        ObjRember.BackColor = TBackup.mlabel(I).nBackColor
        ObjRember.Caption = TBackup.mlabel(I).nCaption
        ObjRember.FontBold = TBackup.mlabel(I).nFontBold
        ObjRember.FontItalic = TBackup.mlabel(I).nFontItalic
        ObjRember.FontUnderline = TBackup.mlabel(I).nFontUnderline
        ObjRember.FontName = TBackup.mlabel(I).nFontName
        ObjRember.FontSize = TBackup.mlabel(I).nFontSize
        ObjRember.Alignment = TBackup.mlabel(I).nAlignment
        ObjRember.AutoSize = TBackup.mlabel(I).nAutoSize
    Next
    I = 0
    For I = 1 To UBound(TBackup.mPictureBox)
        Set ObjRember = Frm.PicImg(I)
        ObjRember.Top = TBackup.mPictureBox(I).nTop
        ObjRember.Left = TBackup.mPictureBox(I).nLeft
        ObjRember.Width = TBackup.mPictureBox(I).nWidth
        ObjRember.Height = TBackup.mPictureBox(I).nHeight
        ObjRember.BackColor = TBackup.mPictureBox(I).nBackColor
        ObjRember.AutoSize = TBackup.mPictureBox(I).nAutoSize
        ObjRember.BorderStyle = TBackup.mPictureBox(I).nBorderStyle
    Next
    I = 0
    For I = 1 To UBound(TBackup.mTextBox)
        Set ObjRember = Frm.txtA(I)
        ObjRember.Top = TBackup.mTextBox(I).nTop
        ObjRember.Left = TBackup.mTextBox(I).nLeft
        ObjRember.Width = TBackup.mTextBox(I).nWidth
        ObjRember.Height = TBackup.mTextBox(I).nHeight
        ObjRember.ForeColor = TBackup.mTextBox(I).nForeColor
        ObjRember.BackColor = TBackup.mTextBox(I).nBackColor
        ObjRember.Text = TBackup.mTextBox(I).nText
        ObjRember.FontBold = TBackup.mTextBox(I).nFontBold
        ObjRember.FontItalic = TBackup.mTextBox(I).nFontItalic
        ObjRember.FontUnderline = TBackup.mTextBox(I).nFontUnderline
        ObjRember.FontName = TBackup.mTextBox(I).nFontName
        ObjRember.FontSize = TBackup.mTextBox(I).nFontSize
    Next
    I = 0
    For I = 1 To UBound(TBackup.mListBox)
        Set ObjRember = Frm.lstA(I)
        ObjRember.Top = TBackup.mListBox(I).nTop
        ObjRember.Left = TBackup.mListBox(I).nLeft
        ObjRember.Width = TBackup.mListBox(I).nWidth
        ObjRember.Height = TBackup.mListBox(I).nHeight
        ObjRember.ForeColor = TBackup.mListBox(I).nForeColor
        ObjRember.BackColor = TBackup.mListBox(I).nBackColor
        ObjRember.FontBold = TBackup.mListBox(I).nFontBold
        ObjRember.FontItalic = TBackup.mListBox(I).nFontItalic
        ObjRember.FontUnderline = TBackup.mListBox(I).nFontUnderline
        ObjRember.FontName = TBackup.mListBox(I).nFontName
        ObjRember.FontSize = TBackup.mListBox(I).nFontSize
        ObjRember.Clear
    Next
    I = 0
    
End Sub
