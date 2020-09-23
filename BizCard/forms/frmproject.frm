VERSION 5.00
Begin VB.Form frmproject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Project"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLan 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1350
      Width           =   1545
   End
   Begin VB.CommandButton cmdfolname 
      Caption         =   "...."
      Height          =   330
      Left            =   5535
      TabIndex        =   2
      Top             =   885
      Width           =   360
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   3
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4815
      TabIndex        =   4
      Top             =   1455
      Width           =   1215
   End
   Begin VB.TextBox txtProLoc 
      Height          =   345
      Left            =   1305
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   878
      Width           =   4125
   End
   Begin VB.TextBox txtProjName 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   360
      Width           =   4635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1365
      Width           =   765
   End
   Begin VB.Label lblProjLoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   945
      Width           =   660
   End
   Begin VB.Label lblProjName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "frmproject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CboTemp As String

Private Sub cboLan_Change()
    cboLan.Text = CboTemp
End Sub

Private Sub cboLan_Click()
    CboTemp = cboLan.Text
End Sub

Private Sub cmdCancel_Click()
    ButtonPressed = 0
    Unload frmproject
End Sub

Private Sub cmdfolname_Click()
Dim folName As String
    folName = GetFolder(frmproject.hwnd, "Folder:")
    
    If Len(folName) = 0 Then Exit Sub
    If Len(folName) = 3 Then txtProLoc.Text = folName: Exit Sub
    txtProLoc.Text = FixPath(folName)
    folName = ""
    
End Sub

Private Sub cmdok_Click()
    ButtonPressed = 1
    ProjectName = txtProjName.Text
    ProjectFolder = FixPath(txtProLoc.Text & txtProjName.Text)
    ProjectFileName = ProjectFolder & ProjectName & ".prj"
    TProject.ProgLan = cboLan.Text
    Unload frmproject
End Sub

Private Sub Form_Load()
    txtProjName.Text = "Project1"
    txtProLoc.Text = FixPath(App.Path) & "Projects\"
    txtProLoc.SelStart = Len(txtProLoc.Text)
    cboLan.AddItem "VBScript"
    cboLan.AddItem "JavaScript"
    cboLan.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmproject = Nothing
End Sub

Private Sub txtProjName_Change()
    cmdok.Enabled = CBool(Len(Trim(txtProjName.Text))) <> False
End Sub
