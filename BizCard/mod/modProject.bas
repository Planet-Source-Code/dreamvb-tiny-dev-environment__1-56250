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
    ProjectTitle As String
    FormFile As String
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
Public ProjectFolder As String
Public ProjectName As String
Public TProject As ProjectFile
Public TUnitSrc As String

Public Sub SaveProject(Frm As Form)
On Error Resume Next
    ' inportant kill the old files as it has a habit of just keeping the old stuff
    Kill TProject.FormFile
    WriteToFile TProject.UnitFile, TUnitSrc
    SaveFrom Frm, TProject.FormFile
    If Err Then Err.Clear
    
End Sub
Private Function GetElement(lzStr As String, nStartPos As String, nEndPos As String) As String
Dim ipos As Integer, lPos As Integer
    ipos = InStr(1, lzStr, nStartPos, vbTextCompare)
    lPos = InStr(ipos + 1, lzStr, nEndPos, vbTextCompare)
    
    If (ipos > 0 And lPos > 0) Then
       GetElement = Mid(lzStr, ipos + Len(nStartPos), lPos - Len(nEndPos) - ipos + 1)
    Else
        GetElement = vbNullChar
    End If
    
    ipos = 0: lPos = 0
End Function

Public Sub CreateProject()
Dim a As String
    ' Build the project file

    
    a = a & "<Project>" & vbCrLf
    a = a & vbTab & "<Title>" & ProjectName & "</Title>" & vbCrLf
    a = a & vbTab & "<Language>" & TProject.ProgLan & "</Language>" & vbCrLf
    a = a & vbTab & "<Form>" & ProjectFolder & ProjectName & ".bfm" & "</Form>" & vbCrLf
    a = a & vbTab & "<unit>" & ProjectFolder & ProjectName & ".unt" & "</unit>" & vbCrLf
    a = a & "</Project>"

    Open ProjectFolder & ProjectName & ".proj" For Binary As #1
        Put #1, , a 'write project data to the file
    Close #1
    a = ""
    
    SaveFormA ProjectFolder & ProjectName & ".bfm"
    
    a = "' ======================================" & vbCrLf
    a = a & "' " & MDIForm1.Caption & "  " & vbCrLf
    a = a & "' ======================================" & vbCrLf
    a = a & "' Project Name: " & ProjectName & vbCrLf
    a = a & "' Date: " & Format(Date, "Short Date") & vbCrLf
    a = a & vbCrLf
    a = a & "' Other comments:" & vbCrLf
    a = a & "' ======================================" & vbCrLf
    a = a & "Sub Main()" & vbCrLf
    a = a & "    'Add your main code here" & vbCrLf
    a = a & "End Sub" & vbCrLf
    a = a & vbCrLf
    a = a & "Call Main()" & vbCrLf
    a = a & vbCrLf
    a = a & "' Add your other controls code here" & vbCrLf
    
    
    Open ProjectFolder & ProjectName & ".unt" For Binary As #1
        Put #1, , a 'write project data to the file
    Close #1
    a = ""
    
End Sub

Public Sub SaveFormA(SaveFile As String)
Dim nFile As Long
    nFile = FreeFile
    
    ReDim Preserve ProjectData.mCommandButton(0)
    ReDim Preserve ProjectData.mlabel(0)
    ReDim Preserve ProjectData.mPictureBox(0)
    ReDim Preserve ProjectData.mTextBox(0)
    ReDim Preserve ProjectData.nListBox(0)
    
    ProjectData.mSIG = "DLG"
    ProjectData.mVersion = 1.1
    ProjectData.mFormData.nBackColor = &H8000000F
    ProjectData.mFormData.nCaption = "Dialog"
    ProjectData.mFormData.nWidth = 4440
    ProjectData.mFormData.nHeight = 3090
    ProjectData.mFormData.nStartPosition = 1
    
    Open SaveFile For Binary As nFile
        Put #nFile, , ProjectData
    Close #nFile
    
    ProjectData.mFormData.nBackColor = 0
    ProjectData.mFormData.nCaption = ""
    ProjectData.mFormData.nHeight = 0
    ProjectData.mFormData.nStartPosition = 0
    ProjectData.mFormData.nWidth = 0
    
End Sub

Public Sub SaveFrom(Frm As Form, SaveFileName As String)
Dim CtrlCounter(4) As Integer, Counter As Long, FrmCtrName As String, ThisObject As Object
Dim nFile As Long
On Error Resume Next
    ProjectData.mSIG = "DLG"
    ProjectData.mVersion = 1.1
    ProjectData.mFormData.nBackColor = Frm.BackColor
    ProjectData.mFormData.nCaption = Frm.Caption
    ProjectData.mFormData.nWidth = Frm.Width
    ProjectData.mFormData.nHeight = Frm.Height
    ProjectData.mFormData.nStartPosition = 1
    
    For Counter = 7 To Frm.Controls.Count - 1
        FrmCtrName = TypeName(Frm.Controls(Counter))
        Set ThisObject = Frm.Controls(Counter)
        Select Case UCase(FrmCtrName)
            Case "COMMANDBUTTON"
                CtrlCounter(0) = CtrlCounter(0) + 1
                ReDim Preserve ProjectData.mCommandButton(CtrlCounter(0))
                ProjectData.mCommandButton(CtrlCounter(0)).nAction = 1
                ProjectData.mCommandButton(CtrlCounter(0)).nActionStr = "C:\"
                ProjectData.mCommandButton(CtrlCounter(0)).nWidth = ThisObject.Width
                ProjectData.mCommandButton(CtrlCounter(0)).nHeight = ThisObject.Height
                ProjectData.mCommandButton(CtrlCounter(0)).nLeft = ThisObject.Left
                ProjectData.mCommandButton(CtrlCounter(0)).nTop = ThisObject.Top
                ProjectData.mCommandButton(CtrlCounter(0)).nBackColor = ThisObject.BackColor
                ProjectData.mCommandButton(CtrlCounter(0)).nCaption = ThisObject.Caption
                ProjectData.mCommandButton(CtrlCounter(0)).nFontBold = ThisObject.FontBold
                ProjectData.mCommandButton(CtrlCounter(0)).nFontItalic = ThisObject.FontItalic
                ProjectData.mCommandButton(CtrlCounter(0)).nFontUnderline = ThisObject.FontUnderline
                ProjectData.mCommandButton(CtrlCounter(0)).nFontName = ThisObject.FontName
                ProjectData.mCommandButton(CtrlCounter(0)).nFontSize = ThisObject.FontSize
            Case "PICTUREBOX"
                CtrlCounter(1) = CtrlCounter(1) + 1
                ReDim Preserve ProjectData.mPictureBox(CtrlCounter(1))
                ProjectData.mPictureBox(CtrlCounter(1)).nAutoSize = ThisObject.AutoSize
                ProjectData.mPictureBox(CtrlCounter(1)).nBackColor = ThisObject.BackColor
                ProjectData.mPictureBox(CtrlCounter(1)).nBorderStyle = ThisObject.BorderStyle
                ProjectData.mPictureBox(CtrlCounter(1)).nWidth = ThisObject.Width
                ProjectData.mPictureBox(CtrlCounter(1)).nHeight = ThisObject.Height
                ProjectData.mPictureBox(CtrlCounter(1)).nTop = ThisObject.Top
                ProjectData.mPictureBox(CtrlCounter(1)).nLeft = ThisObject.Left
            Case "LABEL"
                CtrlCounter(2) = CtrlCounter(2) + 1
                ReDim Preserve ProjectData.mlabel(CtrlCounter(2))
                ProjectData.mlabel(CtrlCounter(2)).nAlignment = ThisObject.Alignment
                ProjectData.mlabel(CtrlCounter(2)).nAutoSize = ThisObject.AutoSize
                ProjectData.mlabel(CtrlCounter(2)).nBackColor = ThisObject.BackColor
                ProjectData.mlabel(CtrlCounter(2)).nCaption = ThisObject.Caption
                ProjectData.mlabel(CtrlCounter(2)).nFontBold = ThisObject.FontBold
                ProjectData.mlabel(CtrlCounter(2)).nFontItalic = ThisObject.FontItalic
                ProjectData.mlabel(CtrlCounter(2)).nFontUnderline = ThisObject.FontUnderline
                ProjectData.mlabel(CtrlCounter(2)).nFontName = ThisObject.FontName
                ProjectData.mlabel(CtrlCounter(2)).nFontSize = ThisObject.FontSize
                ProjectData.mlabel(CtrlCounter(2)).nForeColor = ThisObject.ForeColor
                ProjectData.mlabel(CtrlCounter(2)).nWidth = ThisObject.Width
                ProjectData.mlabel(CtrlCounter(2)).nHeight = ThisObject.Height
                ProjectData.mlabel(CtrlCounter(2)).nTop = ThisObject.Top
                ProjectData.mlabel(CtrlCounter(2)).nLeft = ThisObject.Left
            Case "TEXTBOX"
                CtrlCounter(3) = CtrlCounter(3) + 1
                ReDim Preserve ProjectData.mTextBox(CtrlCounter(3))
                ProjectData.mTextBox(CtrlCounter(3)).nForeColor = ThisObject.ForeColor
                ProjectData.mTextBox(CtrlCounter(3)).nBackColor = ThisObject.BackColor
                ProjectData.mTextBox(CtrlCounter(3)).nFontBold = ThisObject.FontBold
                ProjectData.mTextBox(CtrlCounter(3)).nFontItalic = ThisObject.FontItalic
                ProjectData.mTextBox(CtrlCounter(3)).nFontUnderline = ThisObject.FontUnderline
                ProjectData.mTextBox(CtrlCounter(3)).nFontName = ThisObject.FontName
                ProjectData.mTextBox(CtrlCounter(3)).nFontSize = ThisObject.FontSize
                ProjectData.mTextBox(CtrlCounter(3)).nWidth = ThisObject.Width
                ProjectData.mTextBox(CtrlCounter(3)).nHeight = ThisObject.Height
                ProjectData.mTextBox(CtrlCounter(3)).nTop = ThisObject.Top
                ProjectData.mTextBox(CtrlCounter(3)).nLeft = ThisObject.Left
                ProjectData.mTextBox(CtrlCounter(3)).nText = ThisObject.Text
            Case "LISTBOX"
                CtrlCounter(4) = CtrlCounter(4) + 1
                ReDim Preserve ProjectData.nListBox(CtrlCounter(4))
                ProjectData.nListBox(CtrlCounter(4)).nForeColor = ThisObject.ForeColor
                ProjectData.nListBox(CtrlCounter(4)).nBackColor = ThisObject.BackColor
                ProjectData.nListBox(CtrlCounter(4)).nFontBold = ThisObject.FontBold
                ProjectData.nListBox(CtrlCounter(4)).nFontItalic = ThisObject.FontItalic
                ProjectData.nListBox(CtrlCounter(4)).nFontUnderline = ThisObject.FontUnderline
                ProjectData.nListBox(CtrlCounter(4)).nFontName = ThisObject.FontName
                ProjectData.nListBox(CtrlCounter(4)).nFontSize = ThisObject.FontSize
                ProjectData.nListBox(CtrlCounter(4)).nWidth = ThisObject.Width
                ProjectData.nListBox(CtrlCounter(4)).nHeight = ThisObject.Height
                ProjectData.nListBox(CtrlCounter(4)).nTop = ThisObject.Top
                ProjectData.nListBox(CtrlCounter(4)).nLeft = ThisObject.Left
                
        End Select
    Next

    Set ThisObject = Nothing
    Erase CtrlCounter()
    Counter = 0
    FrmCtrName = ""
    
    'FlushFileBuffers nFile
    nFile = FreeFile
    Open SaveFileName For Binary Access Write As nFile
        Put #nFile, , ProjectData
    Close #nFile
    
   ' FlushFileBuffers nFile
    ProjectData.mFormData.nBackColor = 0
    ProjectData.mFormData.nCaption = ""
    ProjectData.mFormData.nHeight = 0
    ProjectData.mFormData.nStartPosition = 0
    ProjectData.mFormData.nWidth = 0
    
    Erase ProjectData.mCommandButton()
    Erase ProjectData.mlabel()
    Erase ProjectData.mPictureBox()
    Erase ProjectData.mTextBox()
    Erase ProjectData.nListBox()
    
    ReDim Preserve ProjectData.mCommandButton(0)
    ReDim Preserve ProjectData.mPictureBox(0)
    ReDim Preserve ProjectData.mlabel(0)
    ReDim Preserve ProjectData.mTextBox(0)
    ReDim Preserve ProjectData.nListBox(0)
    
    If Err Then AppendErrorLog ("SaveFrom()" & vbCrLf & Err.Description & ";")


End Sub

Public Function LoadForm(Frm As Form, lzFormFile As String) As Boolean
Dim I As Integer, inFile As Long
On Error Resume Next

    If Not IsFileHere(lzFormFile) Then LoadForm = True: Exit Function
    If FileLen(lzFormFile) <= 0 Then LoadForm = True: Exit Function
    
    inFile = FreeFile
    Frm.Visible = False
    
    
    Erase ProjectData.mCommandButton()
    Erase ProjectData.mlabel()
    Erase ProjectData.mlabel()
    Erase ProjectData.mTextBox()
    Erase ProjectData.nListBox()
    
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
    
    If Err Then AppendErrorLog ("LoadForm()" & vbCrLf & Err.Description & ";")

    I = 0
    frmWorkArea.HideSelection
    Frm.Visible = True
    LoadForm = True
End Function

Public Function OpenProject(lzProjectFile As String) As Boolean
Dim StrB As Variant, sText As String

    sText = OpenFile(lzProjectFile)
    StrB = Split(sText, vbCrLf)
    
    If Not LCase(StrB(0)) = "<project>" Then OpenProject = False: Exit Function
    If Not LCase(StrB(UBound(StrB))) = "</project>" Then OpenProject = False: Exit Function


    TProject.ProjectTitle = GetElement(sText, "<title>", "</title>")
    TProject.FormFile = GetElement(sText, "<form>", "</form>")
    TProject.UnitFile = GetElement(sText, "<unit>", "</unit>")
    TProject.ProgLan = GetElement(sText, "<language>", "</language>")
    
    If TProject.ProjectTitle = vbNullChar Then
        OpenProject = False
        Exit Function
    ElseIf TProject.FormFile = vbNullChar Then
        OpenProject = False
        Exit Function
    ElseIf TProject.UnitFile = vbNullChar Then
        OpenProject = False
        Exit Function
    ElseIf TProject.ProgLan = vbNullChar Then
        OpenProject = False
        Exit Function
    Else
        OpenProject = True
    End If
    
    Erase StrB
    sText = ""
End Function
