VERSION 5.00
Begin VB.Form frmWorkArea 
   BackColor       =   &H80000009&
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2685
   ScaleWidth      =   4320
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   0
      Left            =   1380
      TabIndex        =   5
      Top             =   1905
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtA 
      Height          =   525
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Text            =   "TextBox"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1170
   End
   Begin Project1.bSelect Hangle 
      Height          =   90
      Left            =   150
      TabIndex        =   3
      Top             =   2595
      Visible         =   0   'False
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   159
   End
   Begin VB.PictureBox PicImg 
      Height          =   495
      Index           =   0
      Left            =   105
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   930
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "Button"
      Height          =   495
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Selection 
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      Height          =   390
      Left            =   390
      Top             =   2505
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmWorkArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Oldx As Integer, OldY As Integer
Dim CanMove As Boolean, CanResize As Boolean
Dim nTmp As String
Private tForeCol As Long
Public Sub CenterDialog()
    If Not inIde Then
        frmWorkArea.Left = (MDIForm1.ScaleWidth - frmWorkArea.Width) / 2
        frmWorkArea.Top = (MDIForm1.ScaleHeight - frmWorkArea.Height) / 2
    Else
        frmWorkArea.Left = (Screen.Width - frmWorkArea.Width) / 2
        frmWorkArea.Top = (Screen.Height - frmWorkArea.Height) / 2
    End If
    
End Sub
Public Property Get ForeColorf() As Long
    ForeColorf = tForeCol
End Property

Public Property Let ForeColorf(ByVal vNewValue As Long)
    tForeCol = vNewValue
End Property

Public Sub UnloadDialog()
    MsgBox "Please use the stop button in the ide", vbInformation
End Sub
Public Sub HideSelection()
    Selection.Visible = False
    Hangle.Visible = False
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
End Sub

Public Sub MakeSelection(mShow As Boolean)

    Hangle.ZOrder 0
    Selection.Top = (TheObjectName.Top - 65)
    Selection.Left = (TheObjectName.Left - 65)
    
    Selection.Width = (TheObjectName.Width + 130)
    Selection.Height = (TheObjectName.Height + 130)
    
    Hangle.Top = (TheObjectName.Top + Selection.Height)
    Hangle.Left = (TheObjectName.Left + Selection.Width)
    
    Selection.Visible = mShow
    Hangle.Visible = mShow
    
    MDIForm1.StatusBar1.Panels(2).Text = "(X " & TheObjectName.Left & ", Y " & TheObjectName.Top & ")"
    Modified = True ' chnages have been make
    MDIForm1.StatusBar1.Panels(1).Text = "Modified"
    
End Sub

Private Sub TMouseUP(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    If mObject.Left <= 0 Then mObject.Left = 0
    If mObject.Top <= 0 Then mObject.Top = 0
    If (mObject.Left + mObject.Width) >= frmWorkArea.Width Then mObject.Left = (frmWorkArea.Width - mObject.Width)
    If (mObject.Top + mObject.Height) >= frmWorkArea.Height Then mObject.Top = (frmWorkArea.Height - mObject.Width)
    MakeSelection True
    CanMove = False
    ObjectSelected = 1 ' dialog object selected
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
    MDIForm1.Toolbar1.Buttons(15).Enabled = True
    MDIForm1.Toolbar1.Buttons(16).Enabled = True
    MDIForm1.mnucut.Enabled = True
    MDIForm1.mnufront.Enabled = True
    MDIForm1.mnuBack.Enabled = True
    
End Sub
Private Sub TMouseDown(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    Set TheObjectName = mObject
    mObject.ZOrder 0
    Oldx = X
    OldY = Y
    CanMove = True
    MakeSelection True
    
    If Button = vbRightButton Then PopupMenu frmmenu.Mnu1

End Sub
Private Sub MoveControl(mObject As Object, Button As Integer, X As Single, Y As Single)
    If Not inIde Then Exit Sub
    If Button = vbLeftButton Then
        mObject.Left = mObject.Left + (X - Oldx)
        mObject.Top = mObject.Top + (Y - OldY)
        MakeSelection True
        Modified = True
        MDIForm1.StatusBar1.Panels(2).Text = "(X " & mObject.Left & ", Y " & mObject.Top & ")"
    End If
End Sub

Private Sub CmdBut_Click(Index As Integer)
    If inIde Then Exit Sub
    tIndex = Index
    dScript.mRunProc "CmdBut_Click"
End Sub

Private Sub CmdBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown CmdBut(Index), Button, X, Y
End Sub

Private Sub CmdBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl CmdBut(Index), Button, X, Y
End Sub

Private Sub CmdBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP CmdBut(Index), Button, X, Y
End Sub

Private Sub Form_Load()
    frmWorkArea.Width = 4380
    frmWorkArea.Height = 3270
    inIde = True
    ObjectSelected = False
    DrawGrid frmWorkArea
    RemoveControlButton frmWorkArea, 6, False
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideSelection
    MDIForm1.mnucut.Enabled = False
    MDIForm1.mnufront.Enabled = False
    MDIForm1.mnuBack.Enabled = False
    
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
    MDIForm1.Toolbar1.Buttons(15).Enabled = False
    MDIForm1.Toolbar1.Buttons(16).Enabled = False
    
End Sub

Private Sub Form_Resize()
    If Not inIde Then Exit Sub
    DrawGrid frmWorkArea
End Sub

Private Sub Hangle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CanResize = True
        Oldx = X
        OldY = Y
    End If
End Sub

Private Sub Hangle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If (Button = vbLeftButton And CanResize) Then
        Hangle.Top = Hangle.Top - (OldY - Y)
        Hangle.Left = Hangle.Left - (Oldx - X)
        
        Selection.Width = Hangle.Left - (Selection.Left - 8)
        Selection.Height = Hangle.Top - (Selection.Top - 8)
        TheObjectName.Width = Selection.Width - 130
        TheObjectName.Height = Selection.Height - 130
        MDIForm1.StatusBar1.Panels(2).Text = "(W " & TheObjectName.Width & ", H " & TheObjectName.Height & ")"
    End If
End Sub

Private Sub Hangle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    CanResize = False
    TheObjectName.Height = Selection.Height - 130
    TheObjectName.Width = Selection.Width - 130
End Sub

Private Sub lblA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown lblA(Index), Button, X, Y
End Sub

Private Sub lblA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl lblA(Index), Button, X, Y
End Sub

Private Sub lblA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP lblA(Index), Button, X, Y
End Sub

Private Sub lstA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown lstA(Index), Button, X, Y
End Sub

Private Sub lstA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl lstA(Index), Button, X, Y
End Sub

Private Sub lstA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP lstA(Index), Button, X, Y
End Sub

Private Sub PicImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown PicImg(Index), Button, X, Y
End Sub

Private Sub PicImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl PicImg(Index), Button, X, Y
End Sub

Private Sub PicImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP PicImg(Index), Button, X, Y
End Sub

Private Sub txtA_Change(Index As Integer)
    If inIde Then txtA(Index).Text = nTmp
End Sub

Private Sub txtA_Click(Index As Integer)
    If inIde Then nTmp = txtA(Index).Text
End Sub

Private Sub txtA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseDown txtA(Index), Button, X, Y
End Sub

Private Sub txtA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveControl txtA(Index), Button, X, Y
End Sub

Private Sub txtA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TMouseUP txtA(Index), Button, X, Y
End Sub
