VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Tiny - Dev Beta 1"
   ClientHeight    =   5475
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   270
      Top             =   4725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2134
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2486
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5115
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_NEW"
            Object.ToolTipText     =   "New..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_OPEN"
            Object.ToolTipText     =   "Open..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_SAVE"
            Object.ToolTipText     =   "Save.."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_CUT"
            Object.ToolTipText     =   "Cut..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_COPY"
            Object.ToolTipText     =   "Copy..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_PASTE"
            Object.ToolTipText     =   "Paste..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_FIND"
            Object.ToolTipText     =   "Find Text..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_RUN"
            Object.ToolTipText     =   "Run..."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_STOP"
            Object.ToolTipText     =   "Stop..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "T_FORM"
            Object.ToolTipText     =   "View Code..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_BACK"
            Object.ToolTipText     =   "Send To Back..."
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "T_FONT"
            Object.ToolTipText     =   "Bring To Front..."
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnubalnk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save Project"
      End
      Begin VB.Menu mnublank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnumake 
         Caption         =   "&Make Exe"
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnudelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnufront 
         Caption         =   "&Bring to Front"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Send to Back"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnugrid 
         Caption         =   "&Grid"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodeView As Boolean

Sub SetupScript()
    Set dScript.DialogObject = Nothing
    dScript.DialogStrName = ""
    dScript.mLanguage = ""
    dScript.mLanguage = TProject.ProgLan
    dScript.DialogStrName = "Dialog"
    Set dScript.DialogObject = frmWorkArea
    dScript.SetupControl
End Sub
Public Sub IdeStop()
    frmTools.Visible = True ' show the tools menu
    Toolbar1.Buttons(13).Enabled = True ' enable code view button
    inIde = True ' in design mode
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(10).Enabled = True ' enable the run button
    Toolbar1.Buttons(11).Enabled = False ' disable the stop button
    mnunew.Enabled = True
    mnuopen.Enabled = True
    mnusave.Enabled = True
    mnumake.Enabled = True
    mnufront.Enabled = False
    mnuback.Enabled = False
    ' restore the forms data
    RestoreData frmWorkArea
    frmWorkArea.Cls
    Set frmWorkArea.Picture = Nothing
    DrawGrid frmWorkArea
    Set dScript.DialogObject = Nothing
    dScript.Reset
End Sub

Private Sub SetupIDE()
    LoadForm frmWorkArea, TProject.FormFile
    MDIForm1.Caption = "V-Dialog - " & TProject.ProjectTitle
    StatusBar1.Panels(1).Visible = True
    StatusBar1.Panels(2).Visible = True
    frmTools.Visible = True
    frmWorkArea.Visible = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(13).Enabled = True
    mnuedit.Enabled = True
    mnuview.Enabled = True
    frmCode.Visible = False
    mnusave.Enabled = True
    mnumake.Enabled = True
    
    frmCode.txtCode = OpenFile(TProject.UnitFile)
    ' setup the script control
    SetupScript
End Sub
Private Sub UnloadAllControls()
On Error Resume Next
Dim I As Integer
    For I = 1 To frmWorkArea.CmdBut.Count - 1
        Unload frmWorkArea.CmdBut(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.PicImg.Count - 1
        Unload frmWorkArea.PicImg(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.lblA.Count - 1
        Unload frmWorkArea.lblA(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.txtA.Count - 1
        Unload frmWorkArea.txtA(I)
    Next
    I = 0
    For I = 1 To frmWorkArea.lstA.Count - 1
        Unload frmWorkArea.lstA(I)
    Next
    I = 0
    
    frmWorkArea.HideSelection
    If Err Then AppendErrorLog ("Error_UnloadAllControls()" & vbCrLf & Err.Description & ";")
    
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
    FirstTimeLoad = True
    ProjectFolder = FixPath(App.Path) & "Projects\"
    If Not FindDir(ProjectFolder) Then MkDir ProjectFolder
    
    Modified = False
    CodeView = False

    frmWorkArea.Left = (frmTools.Width + frmTools.Left + 40)
    
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(11).Enabled = False
    
    mnucut.Enabled = False
    mnucopy.Enabled = False
    mnupaste.Enabled = False
    mnuall.Enabled = False
    mnudelete.Enabled = False
    mnusave.Enabled = False
    mnumake.Enabled = False
    
    
    ReDim Preserve ProjectData.mCommandButton(0)
    ReDim Preserve ProjectData.mlabel(0)
    ReDim Preserve ProjectData.mPictureBox(0)
    ReDim Preserve ProjectData.mTextBox(0)
    ReDim Preserve ProjectData.nListBox(0)
    frmWorkArea.Hide
    frmTools.Hide
    StatusBar1.Panels(1).Visible = False
    StatusBar1.Panels(2).Visible = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(13).Enabled = False
    mnuedit.Enabled = False
    mnuview.Enabled = False
    
    If Err Then AppendErrorLog ("Error_MDIForm_Load()" & vbCrLf & Err.Description & ";")

End Sub

Private Sub mnuabout_Click()
    FrmAbout.Show vbModal, MDIForm1
End Sub

Private Sub mnuall_Click()
    EditMenu nSelectAll, frmCode.txtCode
End Sub

Private Sub mnuback_Click()
    TheObjectName.ZOrder vbSendToBack
End Sub

Private Sub mnucopy_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nCopy, frmCode.txtCode
    End If
End Sub

Private Sub mnucut_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nCut, frmCode.txtCode
    Else
        frmWorkArea.HideSelection ' hide objects selection
        Unload TheObjectName ' unload the object
    End If
End Sub

Private Sub mnudelete_Click()
    EditMenu nDelete, frmCode.txtCode
End Sub

Private Sub mnufront_Click()
    TheObjectName.ZOrder vbBringToFront
End Sub

Private Sub mnugrid_Click()
Static iCheck As Boolean
    iCheck = Not iCheck
    mnugrid.Checked = Not iCheck
    
    If iCheck Then
        Set frmWorkArea.Picture = Nothing
    Else
        DrawGrid frmWorkArea
    End If
End Sub

Private Sub mnumake_Click()
Dim nDebugPath As String, sBuffer1 As String, sBuffer2 As String, iFile As Long, NewExe As String, _
TheHeadInfo As String, ExeHeadFile As String, sHead As String

    nDebugPath = GetAbsPath(TProject.FormFile) & "debug"
    NewExe = FixPath(nDebugPath) & TProject.ProjectTitle & ".exe"
    ExeHeadFile = FixPath(App.Path) & "exehead\exehead.exe"
    
    If Not IsFileHere(ExeHeadFile) Then
        MsgBox "Compile Error unable to link code." _
        & vbCrLf & vbCrLf & ExeHeadFile & " was not found", vbCritical, "File not Found"
        Exit Sub
    End If

    If IsFileHere(NewExe) Then Kill NewExe
    If Not FindDir(nDebugPath) Then MkDir nDebugPath
    
    mnusave_Click
    
    sBuffer1 = Encrypt(RemoveComments(OpenFile(TProject.UnitFile)))
    sBuffer2 = OpenFile(TProject.FormFile)
    
    MakeExe.Win32CodeData = sBuffer1
    MakeExe.Win32FormData = sBuffer2
    MakeExe.Win32Lan = TProject.ProgLan
    
    sBuffer1 = ""
    sBuffer2 = ""
    
    iFile = FreeFile
    
    FileCopy ExeHeadFile, NewExe
    sHead = "<DATA>" & Chr(5)
    Open NewExe For Binary As #iFile
        Put #iFile, LOF(iFile), sHead
        Put #iFile, LOF(iFile) + 1, MakeExe
    Close #iFile
    
    sHead = ""
    ExeHeadFile = ""
    
    MsgBox "Your Appliaction has now been compiled to :" _
    & vbCrLf & NewExe, vbInformation
    NewExe = ""
    
End Sub

Private Sub mnunew_Click()
    frmproject.Show vbModal, MDIForm1
    If ButtonPressed = 0 Then Exit Sub
    If FindDir(ProjectFolder) Then
        MsgBox "The project name you named already exsits." _
        & vbCrLf & vbCrLf & "Please choose a different name", vbInformation, Me.Caption
        Exit Sub
    Else
        MkDir ProjectFolder
        CreateProject
        If Not OpenProject(ProjectFolder & ProjectName & ".proj") Then
            MsgBox "The project can't be opened", vbCritical, "Unable To Load Project"
            Exit Sub
        Else
            SetupIDE
            FirstTimeLoad = False
        End If
    End If
    
End Sub

Private Sub mnuopen_Click()
    frmOpen.Show vbModal, MDIForm1
    If ButtonPressed = 0 Then Exit Sub
        If Not OpenProject(ProjectFileToOpen) Then
            MsgBox "The project can't be opened", vbCritical, "Unable To Load Project"
            Exit Sub
        Else
        UnloadAllControls
        SetupIDE
    End If
End Sub

Private Sub mnupaste_Click()
    If ObjectSelected = iCodeView Then ' Code view mode
        EditMenu nPaste, frmCode.txtCode
    End If
End Sub

Private Sub mnusave_Click()
    On Error Resume Next
    ' A little bug fix not perfect like but seems to work
    ' had to add the loop below because if you deleted
    ' a Form object eg button and save the old button will still be there.
    ' not sure why it does it seems VB does not like saveing or deleteing files
    ' If you can find a different way please add it in or let me know.
    Dim L As Integer
    For L = 0 To 1
        DeleteFile TProject.FormFile
        DeleteFile TProject.UnitFile
        TUnitSrc = frmCode.txtCode.Text
        SaveProject frmWorkArea
    Next
    L = 0
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

    Select Case UCase(Button.Key)
        Case "T_NEW"    ' Call Menu New
            mnunew_Click
        Case "T_OPEN"   ' Call Menu open
            mnuopen_Click
        Case "T_SAVE"
            mnusave_Click ' Call Menu Save
        Case "T_CUT"
            mnucut_Click ' Call menu Cut
        Case "T_COPY"
            mnucopy_Click ' Call Menu Copy
        Case "T_PASTE"
            mnupaste_Click ' Call menu Paste
    Case "T_BACK"
        mnuback_Click
    Case "T_FONT"
            mnufront_Click
        Case "T_FORM"
                CodeView = Not CodeView
                frmCode.Visible = CodeView
                frmWorkArea.Visible = Not CodeView
                frmTools.Visible = frmWorkArea.Visible
                Toolbar1.Buttons(15).Enabled = Not CodeView
                Toolbar1.Buttons(16).Enabled = Not CodeView
                Toolbar1.Buttons(5).Enabled = False ' disable cut
                Toolbar1.Buttons(6).Enabled = False ' disable copy
                
                If CodeView Then
                    ObjectSelected = 0
                    Toolbar1.Buttons(13).Image = ImageList1.ListImages(10).Index
                    Toolbar1.Buttons(13).ToolTipText = "Form Designer..."
                    Toolbar1.Buttons(7).Enabled = EnablePaste
                    Toolbar1.Buttons(10).Enabled = False
                    mnupaste.Enabled = EnablePaste
                    mnuall.Enabled = True
                Else
                    ObjectSelected = 1
                    Toolbar1.Buttons(13).Image = ImageList1.ListImages(11).Index
                    Toolbar1.Buttons(13).ToolTipText = "View Code..."
                    Toolbar1.Buttons(10).Enabled = True
                    Toolbar1.Buttons(7).Enabled = False ' disable paste button
                    mnupaste.Enabled = False ' disable paste menu item
                    mnuall.Enabled = False
                End If

        Case "T_RUN"
            frmTools.Visible = False ' hide the tools menu
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(13).Enabled = False ' disable code view button
            Toolbar1.Buttons(10).Enabled = False ' disable the run button
            Toolbar1.Buttons(11).Enabled = True ' enable the stop button
            Toolbar1.Buttons(15).Enabled = False
            Toolbar1.Buttons(16).Enabled = False
            mnunew.Enabled = False
            mnuopen.Enabled = False
            mnusave.Enabled = False
            mnumake.Enabled = False
            
            DialogRun frmWorkArea
            ' Remmber the forms data
            RemberFormData frmWorkArea
            ' setup the script control
            SetupScript
            ' run the main code here
            dScript.RunCode frmCode.txtCode.Text
            
        Case "T_STOP"
            IdeStop
    End Select
    
    'If Err Then AppendErrorLog ("Error_Toolbar1_ButtonClick()" & vbCrLf & "Button:" & Button.Key & vbCrLf & Err.Description & ";")

End Sub
