VERSION 5.00
Begin VB.Form FrmFS 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   900
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "FrmFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
FS = False
ChangeRes CSng(OldW), CSng(OldH)
FrmFS.Hide
End Sub

Private Sub Picture1_Click()
FS = False
ChangeRes CSng(OldW), CSng(OldH)
FrmFS.Hide
End Sub
