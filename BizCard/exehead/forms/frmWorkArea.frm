VERSION 5.00
Begin VB.Form frmWorkArea 
   BackColor       =   &H80000009&
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4320
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   0
      ItemData        =   "frmWorkArea.frx":0000
      Left            =   1290
      List            =   "frmWorkArea.frx":0002
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtA 
      Height          =   525
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Text            =   "TextBox"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1170
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
Private tForeCol As Long

Public Property Get ForeColorf() As Long
    ForeColorf = tForeCol
End Property

Public Property Let ForeColorf(ByVal vNewValue As Long)
    tForeCol = vNewValue
End Property

Public Sub UnloadDialog()
    CleanUpAll
    Unload frmWorkArea
End Sub

Public Sub CenterDialog()
    frmWorkArea.Left = (Screen.Width - frmWorkArea.Width) / 2
    frmWorkArea.Top = (Screen.Height - frmWorkArea.Height) / 2
End Sub

Private Sub CmdBut_Click(Index As Integer)
    tIndex = Index
    dScript.mRunProc "CmdBut_Click"
End Sub

Private Sub Form_Load()
    frmWorkArea.Width = 4380
    frmWorkArea.Height = 3270
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmWorkArea = Nothing
    End
End Sub
