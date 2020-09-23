VERSION 5.00
Begin VB.Form frmmenu 
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   2790
   Begin VB.Menu Mnu1 
      Caption         =   "Top"
      Begin VB.Menu mnudeleteA 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnufront 
         Caption         =   "&Bring to Front"
      End
      Begin VB.Menu mnuback 
         Caption         =   "&Send to Back"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprop 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuback_Click()
    TheObjectName.ZOrder vbSendToBack ' send object to the back
End Sub

Private Sub mnudeleteA_Click()
    frmWorkArea.HideSelection ' hide objects selection
    Unload TheObjectName ' unload the object
End Sub

Private Sub mnufront_Click()
    TheObjectName.ZOrder vbBringToFront ' bring object to the front
End Sub
