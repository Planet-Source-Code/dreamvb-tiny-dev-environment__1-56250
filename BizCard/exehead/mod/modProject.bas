Attribute VB_Name = "modProject"
Private Type TForm
    nCaption As String
    nHeight As Long
    nWidth As Long
    nBackColor As Long
    nStartPosition As Integer
End Type

Private Type TCommandButton
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nCaption As String
    nAction As Integer
    nActionStr As String
    nBackColor As Long
    nFontName As String
    nFontSize As Long
    nFontBold As Boolean
    nFontItalic As Boolean
    nFontUnderline As Boolean
End Type

Private Type TPicture
    nTop As Long
    nLeft As Long
    nWidth As Long
    nHeight As Long
    nBackColor As Long
    nBorderStyle As Integer
    'nPicture As picture
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

Private Type TLabel
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

Private Type TTextBox
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

Private Type TDialog
    mSIG As String * 3 '"DLG
    mVersion As Single
    mFormData As TForm
    mCommandButton() As TCommandButton
    mPictureBox() As TPicture
    mlabel() As TLabel
    mTextBox() As TTextBox
    nListBox() As TListBox
End Type

Private Type ProjectFile
    UnitFile As String
    ProgLan As String
End Type

Private Type Win32App
    Win32FormData As String
    Win32CodeData As String
    Win32Lan As String
End Type

Public MakeExe As Win32App
Public ProjectData As TDialog
Public TProject As ProjectFile

Public Function LoadForm(Frm As Form, lzFormFile As String) As Boolean
Dim I As Integer, inFile As Long
On Error Resume Next

    If Not IsFileHere(lzFormFile) Then LoadForm = True: Exit Function
    If FileLen(lzFormFile) <= 0 Then LoadForm = True: Exit Function
    
    inFile = FreeFile
    'Frm.Visible = False
    
    Erase ProjectData.mCommandButton()
    Erase ProjectData.mlabel()
    Erase ProjectData.mlabel()
    Erase ProjectData.mTextBox()

    Open lzFormFile For Binary As #inFile
        Get #inFile, , ProjectData
    Close #inFile

    If ProjectData.mSIG <> "DLG" Then LoadForm = False: Exit Function
    If ProjectData.mVersion < 1.1 Then LoadForm = False: Exit Function
    
    Frm.Caption = ProjectData.mFormData.nCaption
    Frm.BackColor = ProjectData.mFormData.nBackColor
    Frm.Width = ProjectData.mFormData.nWidth
    Frm.Height = ProjectData.mFormData.nHeight

    For I = 1 To UBound(ProjectData.mCommandButton)
        tAddControl Frm, "T_BUTTON"
        TheObjectName.Top = ProjectData.mCommandButton(I).nTop
        TheObjectName.Left = ProjectData.mCommandButton(I).nLeft
        TheObjectName.Width = ProjectData.mCommandButton(I).nWidth
        TheObjectName.Height = ProjectData.mCommandButton(I).nHeight
        TheObjectName.BackColor = ProjectData.mCommandButton(I).nBackColor
        TheObjectName.Caption = ProjectData.mCommandButton(I).nCaption
        TheObjectName.FontBold = ProjectData.mCommandButton(I).nFontBold
        TheObjectName.FontItalic = ProjectData.mCommandButton(I).nFontItalic
        TheObjectName.FontUnderline = ProjectData.mCommandButton(I).nFontUnderline
        TheObjectName.FontName = ProjectData.mCommandButton(I).nFontName
        TheObjectName.FontSize = ProjectData.mCommandButton(I).nFontSize
    Next
    I = 0
    For I = 1 To UBound(ProjectData.mlabel)
        tAddControl Frm, "T_LABEL"
        TheObjectName.Top = ProjectData.mlabel(I).nTop
        TheObjectName.Left = ProjectData.mlabel(I).nLeft
        TheObjectName.Width = ProjectData.mlabel(I).nWidth
        TheObjectName.Height = ProjectData.mlabel(I).nHeight
        TheObjectName.ForeColor = ProjectData.mlabel(I).nForeColor
        TheObjectName.BackColor = ProjectData.mlabel(I).nBackColor
        TheObjectName.Caption = ProjectData.mlabel(I).nCaption
        TheObjectName.FontBold = ProjectData.mlabel(I).nFontBold
        TheObjectName.FontItalic = ProjectData.mlabel(I).nFontItalic
        TheObjectName.FontUnderline = ProjectData.mlabel(I).nFontUnderline
        TheObjectName.FontName = ProjectData.mlabel(I).nFontName
        TheObjectName.FontSize = ProjectData.mlabel(I).nFontSize
        TheObjectName.Alignment = ProjectData.mlabel(I).nAlignment
        TheObjectName.AutoSize = ProjectData.mlabel(I).nAutoSize
    Next
    I = 0
    For I = 1 To UBound(ProjectData.mPictureBox)
        tAddControl Frm, "T_IMAGE"
        TheObjectName.Top = ProjectData.mPictureBox(I).nTop
        TheObjectName.Left = ProjectData.mPictureBox(I).nLeft
        TheObjectName.Width = ProjectData.mPictureBox(I).nWidth
        TheObjectName.Height = ProjectData.mPictureBox(I).nHeight
        TheObjectName.BackColor = ProjectData.mPictureBox(I).nBackColor
        TheObjectName.AutoSize = ProjectData.mPictureBox(I).nAutoSize
        TheObjectName.BorderStyle = ProjectData.mPictureBox(I).nBorderStyle
    Next
    I = 0
    For I = 1 To UBound(ProjectData.mTextBox)
        tAddControl Frm, "T_TEXT"
        TheObjectName.Top = ProjectData.mTextBox(I).nTop
        TheObjectName.Left = ProjectData.mTextBox(I).nLeft
        TheObjectName.Width = ProjectData.mTextBox(I).nWidth
        TheObjectName.Height = ProjectData.mTextBox(I).nHeight
        TheObjectName.ForeColor = ProjectData.mTextBox(I).nForeColor
        TheObjectName.BackColor = ProjectData.mTextBox(I).nBackColor
        TheObjectName.Text = ProjectData.mTextBox(I).nText
        TheObjectName.FontBold = ProjectData.mTextBox(I).nFontBold
        TheObjectName.FontItalic = ProjectData.mTextBox(I).nFontItalic
        TheObjectName.FontUnderline = ProjectData.mTextBox(I).nFontUnderline
        TheObjectName.FontName = ProjectData.mTextBox(I).nFontName
        TheObjectName.FontSize = ProjectData.mTextBox(I).nFontSize
    Next
    I = 0
    For I = 1 To UBound(ProjectData.nListBox)
        tAddControl Frm, "T_LIST"
        TheObjectName.Top = ProjectData.nListBox(I).nTop
        TheObjectName.Left = ProjectData.nListBox(I).nLeft
        TheObjectName.Width = ProjectData.nListBox(I).nWidth
        TheObjectName.Height = ProjectData.nListBox(I).nHeight
        TheObjectName.ForeColor = ProjectData.nListBox(I).nForeColor
        TheObjectName.BackColor = ProjectData.nListBox(I).nBackColor
        TheObjectName.FontBold = ProjectData.nListBox(I).nFontBold
        TheObjectName.FontItalic = ProjectData.nListBox(I).nFontItalic
        TheObjectName.FontUnderline = ProjectData.nListBox(I).nFontUnderline
        TheObjectName.FontName = ProjectData.nListBox(I).nFontName
        TheObjectName.FontSize = ProjectData.nListBox(I).nFontSize
    Next
    I = 0
    LoadForm = True
End Function
