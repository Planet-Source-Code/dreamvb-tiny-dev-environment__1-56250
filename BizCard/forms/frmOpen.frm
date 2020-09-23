VERSION 5.00
Begin VB.Form frmOpen 
   Caption         =   "Open Project"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3240
      TabIndex        =   4
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4575
      TabIndex        =   3
      Top             =   2970
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2925
      Pattern         =   "*.proj"
      TabIndex        =   2
      Top             =   150
      Width           =   2880
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCan_Click()
    ButtonPressed = 0
    Unload frmOpen
End Sub

Private Sub cmdOpen_Click()
     ButtonPressed = 1
     ProjectFileToOpen = FixPath(File1.Path) & File1.Filename
     Unload frmOpen
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
    If Err Then MsgBox Err.Description, vbInformation, "Error " & Err.Number
End Sub

Private Sub File1_Click()
    cmdOpen.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpen = Nothing
    
End Sub
