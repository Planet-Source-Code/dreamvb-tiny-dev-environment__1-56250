VERSION 5.00
Begin VB.Form frmCode 
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   5475
   Begin VB.ComboBox CboList 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   2565
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   -15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   2445
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cboTmp As String

Function EnableCutCopy() As Boolean
    If Len(txtCode.SelText) = 0 Then EnableCutCopy = False: Exit Function
    EnableCutCopy = True
End Function

Private Sub CboList_Change()
 CboList.Text = cboTmp
End Sub

Private Sub CboList_Click()
    cboTmp = CboList.Text
End Sub

Private Sub Form_Activate()
    MDIForm1.Toolbar1.Buttons(5).Enabled = EnableCutCopy
    MDIForm1.Toolbar1.Buttons(6).Enabled = EnableCutCopy
    lstVBFunctions txtCode.Text, CboList
    RemoveControlButton frmCode, 6, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cboTmp = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
    txtCode.Width = (frmCode.ScaleWidth)
    txtCode.Height = (frmCode.ScaleHeight - CboList.Height - 50)
End Sub

Private Sub txtCode_Change()
    Modified = True ' chnages have been make
    MDIForm1.StatusBar1.Panels(1).Text = "Modified"
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        txtCode.SelText = Space(4)
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then lstVBFunctions txtCode.Text, CboList
End Sub

Private Sub txtCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDIForm1.Toolbar1.Buttons(5).Enabled = EnableCutCopy
    MDIForm1.Toolbar1.Buttons(6).Enabled = EnableCutCopy
    MDIForm1.mnucut.Enabled = EnableCutCopy
    MDIForm1.mnucopy.Enabled = EnableCutCopy
    MDIForm1.mnudelete.Enabled = EnableCutCopy
    lstVBFunctions txtCode.Text, CboList
End Sub
