VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Animation"
   ClientHeight    =   1395
   ClientLeft      =   2655
   ClientTop       =   2850
   ClientWidth     =   4575
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Reset Particles"
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2490
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1455
      TabIndex        =   5
      Text            =   "100"
      Top             =   765
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3345
      TabIndex        =   2
      Top             =   300
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   330
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   3330
      TabIndex        =   0
      Top             =   795
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Number of frames:"
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   795
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   "Save Path:"
      Height          =   225
      Left            =   75
      TabIndex        =   3
      Top             =   105
      Width           =   3180
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim I As Long
frmMain.Saving = True
frmMain.SDest = Text1.Text
frmMain.SNum = Text2.Text

If Check1.Value = 1 Then
    For I = 0 To 10000
    P(I).Life = 0
    Next
End If

frmSave.Hide
frmMain.Command3_Click
End Sub

Private Sub Command2_Click()
Text1.Text = BrowseForFolder(frmSave.hwnd, "Choose a save path:")

End Sub

Private Sub Text2_Change()
Text2.Text = Val(Text2.Text)
End Sub
