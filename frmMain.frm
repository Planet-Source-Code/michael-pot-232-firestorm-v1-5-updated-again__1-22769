VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firestorm v1.5"
   ClientHeight    =   6555
   ClientLeft      =   2340
   ClientTop       =   1860
   ClientWidth     =   6270
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3165
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   90
      TabIndex        =   47
      Top             =   3165
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fullscreen"
            Object.ToolTipText     =   "Fullscreen"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3210
      Top             =   6345
   End
   Begin VB.PictureBox picOrg 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   3645
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   42
      Top             =   6390
      Visible         =   0   'False
      Width           =   3000
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   240
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1EB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":224E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2986
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Help"
      Height          =   390
      Left            =   2145
      TabIndex        =   41
      Top             =   6600
      Width           =   915
   End
   Begin VB.CommandButton Command11 
      Caption         =   "About"
      Height          =   390
      Left            =   1125
      TabIndex        =   40
      Top             =   6600
      Width           =   930
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   390
      Left            =   120
      TabIndex        =   39
      Top             =   6600
      Width           =   930
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presets"
      Height          =   795
      Left            =   60
      TabIndex        =   1
      Top             =   4905
      Width           =   3030
      Begin VB.CommandButton Command5 
         Caption         =   "Add a preset"
         Height          =   420
         Left            =   105
         TabIndex        =   3
         Top             =   255
         Width           =   1245
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1425
         TabIndex        =   2
         Text            =   "Presets"
         Top             =   300
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Animation Capture"
      Height          =   765
      Left            =   60
      TabIndex        =   30
      Top             =   5730
      Width           =   3030
      Begin VB.CommandButton Command4 
         Caption         =   "Save "
         Height          =   390
         Left            =   120
         TabIndex        =   31
         Top             =   255
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Idle"
         Height          =   285
         Left            =   1245
         TabIndex        =   32
         Top             =   330
         Width           =   1665
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Timeline"
      Enabled         =   0   'False
      Height          =   1050
      Left            =   15
      TabIndex        =   29
      Top             =   7215
      Width           =   6195
      Begin VB.CommandButton Command10 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2310
         TabIndex        =   38
         Top             =   690
         Width           =   285
      End
      Begin VB.CommandButton Command9 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         TabIndex        =   37
         Top             =   690
         Width           =   270
      End
      Begin VB.CommandButton Command8 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1635
         TabIndex        =   36
         Top             =   690
         Width           =   285
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add Keyframe"
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   675
         Width           =   1140
      End
      Begin VB.PictureBox Picture3 
         Height          =   360
         Left            =   120
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   393
         TabIndex        =   33
         Top             =   285
         Width           =   5955
         Begin VB.CommandButton Command6 
            Caption         =   "< >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   90
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   60
      Width           =   3000
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info"
      Height          =   555
      Left            =   90
      TabIndex        =   26
      Top             =   3600
      Width           =   3000
      Begin VB.Label lblBits 
         AutoSize        =   -1  'True
         Caption         =   "Bits per pixel:"
         Height          =   195
         Left            =   1380
         TabIndex        =   28
         Top             =   255
         Width           =   930
      End
      Begin VB.Label lblFPS 
         AutoSize        =   -1  'True
         Caption         =   "FPS:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Palettes"
      Height          =   675
      Left            =   75
      TabIndex        =   23
      Top             =   4185
      Width           =   3015
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1395
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "Palettes"
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Label7 
         Caption         =   "Current Palette:"
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   315
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   6480
      Left            =   3195
      TabIndex        =   4
      Top             =   15
      Width           =   3030
      Begin VB.TextBox Text 
         Height          =   300
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Text            =   "100"
         Top             =   5010
         Width           =   660
      End
      Begin VB.TextBox Text 
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Text            =   "5"
         Top             =   5010
         Width           =   660
      End
      Begin VB.TextBox Text 
         Height          =   300
         Index           =   2
         Left            =   1830
         TabIndex        =   5
         Text            =   "1"
         Top             =   5010
         Width           =   660
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   465
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Max             =   60
         SelStart        =   2
         TickFrequency   =   3
         Value           =   2
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1215
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -60
         Max             =   60
         SelStart        =   4
         TickFrequency   =   3
         Value           =   4
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2025
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Max             =   10000
         SelStart        =   1000
         TickFrequency   =   500
         Value           =   1000
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2775
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Max             =   200
         SelStart        =   4
         TickFrequency   =   10
         Value           =   4
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   3525
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Max             =   255
         SelStart        =   102
         TickFrequency   =   8
         Value           =   102
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   4245
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -1000
         Max             =   1000
         TickFrequency   =   50
      End
      Begin VB.Frame Frame7 
         Height          =   945
         Left            =   105
         TabIndex        =   43
         Top             =   5445
         Width           =   2820
         Begin VB.CheckBox Check2 
            Caption         =   "Bit-mapped Origin"
            Height          =   405
            Left            =   165
            TabIndex        =   46
            Top             =   -75
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Clear Origin"
            Enabled         =   0   'False
            Height          =   390
            Left            =   90
            TabIndex        =   45
            Top             =   405
            Width           =   1275
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Load Origin"
            Enabled         =   0   'False
            Height          =   390
            Left            =   1425
            TabIndex        =   44
            Top             =   405
            Width           =   1275
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Lateral Deviation:"
         Height          =   390
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   2625
      End
      Begin VB.Label Label2 
         Caption         =   "Force:"
         Height          =   390
         Left            =   120
         TabIndex        =   21
         Top             =   1035
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   "Density:"
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Top             =   1830
         Width           =   2625
      End
      Begin VB.Label Label4 
         Caption         =   "Base width:"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         Top             =   2580
         Width           =   2625
      End
      Begin VB.Label Label5 
         Caption         =   "Length:"
         Height          =   390
         Left            =   120
         TabIndex        =   18
         Top             =   3330
         Width           =   2625
      End
      Begin VB.Label Label6 
         Caption         =   "Gravity:"
         Height          =   390
         Left            =   135
         TabIndex        =   17
         Top             =   4050
         Width           =   2625
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   4785
         Width           =   150
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1020
         TabIndex        =   15
         Top             =   4785
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Blur Iterations:"
         Height          =   195
         Left            =   1830
         TabIndex        =   14
         Top             =   4785
         Width           =   1005
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAddPreset 
         Caption         =   "Add as preset"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save as animation"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuCtrl 
      Caption         =   "Control"
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClrOrig 
         Caption         =   "Clear Origin"
      End
      Begin VB.Menu mnuLdOrig 
         Caption         =   "Load Origin"
      End
      Begin VB.Menu mnuRSOrig 
         Caption         =   "Rescan Origin"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuFSHelp 
         Caption         =   "Firestorm Help"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BPF As Long, Hgt As Long, BPP As Long, C As Single
Private B(0 To 40400) As Byte, B2(0 To 40400) As Byte, X As Long, Y As Long, D As Long
Public Ending As Boolean, Pos As Long, I As Long, J As Long, CC As Long, II As Long, Ended As Boolean
Public MDown As Boolean
Private Pal(0 To 255) As RGBQUAD
Public Saving As Boolean, SNum As Long, SDest As String, SCur As Long

Private Type Preset
Name As String
PS(1 To 12) As Long
End Type

Private Type PresetList
PreCount As Long
Pre() As Preset
End Type

Public Pause As Boolean
Private Col As Long, VV As Single, SVV As Single, Svv2 As Single, S3 As Long, S4 As Long, Sze As Long, Heat As Single, Grav As Single, sx As Long, sy As Long, Bl As Long, T As Long, Tim
Private Tot As Single
Private Pr As PresetList

Private Sub Check2_Click()
If Check2.Value = 1 And OCnt < 10 Then
Command1_Click
End If
Command1.Enabled = Check2.Value = 1
Command2.Enabled = Check2.Value = 1
End Sub

Private Sub Combo1_Click()
For I = 0 To 10000
P(I).Life = Rnd * 5
Next
ApplyPreset (Combo1.ListIndex + 1)
End Sub

Private Sub Combo2_Click()
LoadPal Combo2.List(Combo2.ListIndex)
End Sub

Private Sub LoadPal(F As String)
Dim F2 As String

Open App.Path & "\Palettes\" & F & ".pal" For Binary As #1
Get #1, , Pal
Close #1

F2 = App.Path & "\Palettes\" & F & ".tmp"

SavePalAsBitmap F2, Pal()

Picture1.Picture = LoadPicture(F2)
DoEvents

Kill F2

BPP = GetBPP(Picture1)
lblBits.Caption = "Bits per pixel: " & BPP
BPP = BPP / 8
If BPP = 0 Then BPP = 1

Resu:
Pause = False
End Sub

Private Sub Command1_Click()
ReDim O(0 To 10) As BitmapOrigin
Let OCnt = 10
For I = 0 To 10
O(I).Y = -100
Next
End Sub

Private Sub Command11_Click()
MsgBox "Firestorm V1.5 (C) 2001" & vbCr & "By Michael Pote" & vbCr & "michaelpote@worldonline.co.za" & vbCr & vbCr & "Want the VB source? Mail Me", vbInformation
End Sub

Private Sub Command12_Click()
Command3_Click
ShellExecute frmMain.hwnd, "Open", App.Path & "\Help.htm", "", App.Path, 0
End Sub

Private Sub Command2_Click()
ScanOrigin
End Sub

Public Sub Command3_Click()
Pause = (Pause = False)
Toolbar1.Buttons(1).Value = IIf(Pause = True, 1, 0)
mnuPause.Checked = Pause
If Pause Then Command3.Caption = "Resume" Else Command3.Caption = "Pause"
End Sub

Sub SetVars()
VV = Slider(1).Value / 8
SVV = Slider(0).Value / 8
Svv2 = Slider(0).Value / 16
S3 = Slider(3).Value
S4 = Slider(3).Value / 2
Heat = Slider(4).Value / 200
Grav = Slider(5).Value / 1000
sx = Val(Text(0).Text)
sy = Val(Text(1).Text)

Bl = Val(Text(2).Text)

Label1.Caption = "Deviation: " & Slider(0).Value
Label2.Caption = "Force: " & Slider(1).Value
Label3.Caption = "Density: " & Slider(2).Value
Label4.Caption = "Base width: " & Slider(3).Value
Label5.Caption = "Length: " & Slider(4).Value
Label6.Caption = "Gravity: " & Slider(5).Value
End Sub

Private Sub Command4_Click()
Pause = False
Command3_Click
frmSave.Show
End Sub

Private Sub Command5_Click()
AddPreset
End Sub

Private Sub Command8_Click()
If Command8.Caption = "4" Then Command8.Caption = ";" Else Command8.Caption = "4"

End Sub

Private Sub Form_Load()
' FIRESTORM V1.5
' --------------
'
' New Features in V1.5:
' ---------------------
' New palette format (40 times smaller than previous)
' Compile optimisations: 65 FPS!! (See Note below)
' FULL Help file!
'
' NOTE: Compile Optimisations:
' ----------------------------
' Turning off Array Bounds check in the Advanced Optimizations
' Menu when compliling an EXE will push the frame rate up by
' 20 FPS. However, The program WILL become unstable and have an
' error when you quit and sometimes when playing with large amounts
' of particles. I have ironed out many bugs which made it crash but
' I just cant get rid of the exiting error, Otherwise its generaly
' stable.
'
'
' New Features in V1.3:
' --------------------
' New GUI layout, Easier palette access, Animation saving
' 2 New Palettes, Bitmap Origin (see below)
'
' Bitmap Origin
' -------------
' Bit-mapped origin is a new feature in firestorm, which allows
' you to control the positioning of the fire more accuratly.
' All you do is load up a palette file in a paint program,
' Then select a color from the palette with a palette index
' of 250 or higher. then draw some text it that color.
' save it as a Bmp file in the palettes directory, and run
' fire storm. Load your palette and select Bit-mapped origin
' you will see that the fire is being emitted by the text!
' pretty cool, huh?
'
'This code is intended for show only, and was not expected to teach
'anyone anything, thus the lack of comments...

'To use: 1. Compile into an EXE otherwise it goes too slow.
'        2. Select a palette
'        3. Tweak the settings until you're happy.
'        4. Sit back and enjoy
'        5. Improve it!

OldW = Screen.Width / Screen.TwipsPerPixelX
OldH = Screen.Height / Screen.TwipsPerPixelY

Show
MainLoop
End Sub

Sub MainLoop()
On Error Resume Next


ScanForPals

BPP = 1

BPF = Picture1.Width * BPP
Hgt = Picture1.Height

LoadPresets

SetVars

Tim = Timer

Do 'BEGIN MAIN LOOP
DoEvents
    For Y = Hgt - 1 To Hgt + 2 'Clear top and bottom pixels.
    For X = 1 To BPF
    B((BPF * Y) + X) = 0
    Next
    Next
    
If Pause = False Then
    
S3 = Slider(3).Value + ((Rnd * 10) - 5)
S4 = (Slider(3).Value / 2) + ((Rnd * 10) - 5)


If Check2.Value = 0 Then 'BITMAP ORIGIN

    For I = 0 To Slider(2).Value 'BEGIN PARTICLE ENGINE
    With P(I)
    
    If .Life <= 0 Then
    .SV = (Rnd * SVV) - Svv2
    .X = sx + Int(Rnd * S3) - S4
    .Y = sy
    .V = Rnd * VV
    .Life = (Rnd * 155 + 100) * Heat
    If .Life < 0 Then .Life = 0
    End If

    .X = .X + .SV
    .Y = .Y + .V
    .V = .V - Grav
    If .X <= 0 Or .X >= 200 Then .Life = 1
    If .Y <= 1 Or .Y >= 198 Then .Y = 0: .Life = 1
    
    
    Pos = ((200 - Int(.Y)) * BPF) + Int(.X)
    Col = B(Pos) * 2 + (.Life * 2)
    If Col <= 10 Then Col = 0
    If Col > 255 Then Col = 255
    B(Pos) = Col
    
    .Life = .Life - 1
    
    End With
    Next 'END PARTICLE ENGINE
    
ElseIf OCnt > 0 Then

    For I = 0 To Slider(2).Value 'BEGIN PARTICLE ENGINE
    With P(I)
    
    If .Life <= 0 Then
    .SV = (Rnd * SVV) - Svv2
    OIn = Int(Rnd * (OCnt - 1))
    If OIn < 0 Then OIn = 0: MsgBox "Caught a particle below 0"
    If OIn > OCnt Then OIn = UBound(O): MsgBox "Caught a paritcle Abv " & OCnt
    .X = O(OIn).X
    .Y = O(OIn).Y
    .V = Rnd * VV
    .Life = (Rnd * 155 + 100) * Heat
    If .Life < 0 Then .Life = 0
    End If

    .X = .X + .SV
    .Y = .Y + .V
    .V = .V - Grav
    If .X <= 0 Or .X >= 200 Then .Life = 0: GoTo SKK
    If .Y <= 1 Or .Y >= 198 Then .Y = 1: .Life = 0: GoTo SKK
    
    
    Pos = ((200 - Int(.Y)) * BPF) + Int(.X)
    Col = B(Pos) * 2 + (.Life * 2)
    If Col <= 10 Then Col = 0
    If Col > 255 Then Col = 255
    B(Pos) = Col
    
    .Life = .Life - 1
    
SKK:
    
    End With
    Next 'END PARTICLE ENGINE

End If
    
    
    For I = 1 To Bl
    For X = 0 To 200
    For Y = 0 To 200
    Pos = (Y * BPF) + X
    Tot = (CInt(B(Pos - BPF)) + B(Pos + BPF) + B(Pos - BPP) + B(Pos + BPP) + B(Pos + BPP + BPF) + B(Pos - BPP + BPF) + B(Pos + BPP - BPF) + B(Pos - BPP - BPF) + B(Pos))
    If Tot > 1 Then Tot = Tot / 9.51
    B(Pos - BPF) = Tot
    Next
    Next
    Next 'END BLUR FILTERING
    
 
    
    
    SetBitmapBits Picture1.Picture, UBound(B), B(1)
    
    Picture1.Refresh

    If FS Then 'Fullscreen
    FrmFS.Picture1.Picture = Picture1.Picture
    End If

End If 'end of Pause block

    
    T = T + 1 'FPS COUNTER
    If Timer >= Tim + 1 Then
    lblFPS = "FPS: " & T
    Tim = Timer
    T = 0
    End If


If Saving Then 'ANIMATION SAVING

SavePicture Picture1.Picture, SDest & "\FS" & Format(SCur, "0000") & ".bmp"

Label11.Caption = "Saving file " & SCur

If SCur = SNum Then Saving = False: Label11.Caption = "Saving Complete"
SCur = SCur + 1
End If


Loop Until Ending = True
Ended = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Ending = True
SavePresets
Timer1.Enabled = True
Cancel = 1

End Sub

Private Sub mnuAbout_Click()
Command11_Click
End Sub

Private Sub mnuAddPreset_Click()
AddPreset
End Sub

Private Sub mnuClrOrig_Click()
Command1_Click
End Sub

Private Sub mnuEnd_Click()
Form_QueryUnload 1, 0
End Sub

Private Sub mnuFSHelp_Click()
Command12_Click
End Sub

Private Sub mnuLdOrig_Click()
Command2_Click
End Sub

Private Sub mnuPause_Click()
Command3_Click
End Sub

Private Sub mnuSave_Click()
Command4_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If MDown Then
Dim Mx As Long, My As Long
OCnt = OCnt + 4
ReDim Preserve O(0 To OCnt) As BitmapOrigin
DoEvents


O(OCnt).X = X
O(OCnt).Y = 200 - Y
O(OCnt - 1).X = X - 1
O(OCnt - 1).Y = 200 - Y - 1
O(OCnt - 2).X = X + 1
O(OCnt - 2).Y = 200 - Y - 1
O(OCnt - 3).X = X - 1
O(OCnt - 3).Y = 200 - Y + 1
End If

If Y > 199 Then Y = 198
If Y < 1 Then Y = 2

CC = 255
If Saving Or Pause Then GoTo Nxt
For J = -2 To 2
For I = -2 To 2
If Y + J > 200 Then GoTo Skip
If Y - J < 0 Then GoTo Skip
If X + I > 200 Then GoTo Skip
If X - I < 0 Then GoTo Skip

'Paint pixels to picture under mouse cursor

B(((Y + J) * BPF) + (X + I)) = CC
B(((Y - J) * BPF) + (X - I)) = CC
Skip:
Next
Next

Nxt:
End Sub

Sub LoadPresets()
Dim Xw As Long
If Dir(App.Path & "\Presets.dat") = "" Then Exit Sub

Open App.Path & "\Presets.dat" For Binary As #1
Get #1, , Pr
Close #1

For Xw = 1 To Pr.PreCount
Combo1.AddItem Pr.Pre(Xw).Name
Next


End Sub

Sub ApplyPreset(Ind As Long)
With Pr.Pre(Ind)

For I = 0 To 5
Slider(I).Value = .PS(I + 1)
Next

For I = 0 To 2
Text(I).Text = CStr(.PS(I + 7))
Next

Check2.Value = .PS(11)
Combo2.ListIndex = .PS(12)

End With
End Sub

Sub AddPreset()
Dim N As String
N = InputBox("Type in a name for this preset:")
If N = "" Then Exit Sub
Pr.PreCount = Pr.PreCount + 1
ReDim Preserve Pr.Pre(0 To Pr.PreCount) As Preset
With Pr.Pre(Pr.PreCount)

For I = 0 To 5
.PS(I + 1) = Slider(I).Value
Next

For I = 7 To 9
.PS(I) = Val(Text(I - 7).Text)
Next

'.PS(10) = Check1.Value 'Random shift
.PS(11) = Check2.Value 'Bitmap origin
.PS(12) = Combo2.ListIndex 'Palette

.Name = N
Combo1.AddItem .Name
End With
End Sub

Sub SavePresets()
Open App.Path & "\Presets.dat" For Binary As #1
Put #1, , Pr
Close #1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDown = False
End Sub

Private Sub Slider_Change(Index As Integer)
SetVars
End Sub

Private Sub Slider_Scroll(Index As Integer)
SetVars
End Sub

Sub ScanOrigin()
Dim F As String, Bb As BITMAP
F = OpenDialog(frmMain, "Bitmaps|*.bmp", "Open a 200x200 8-bit Bitmap", "")
If F = "" Then Exit Sub

picOrg.Picture = LoadPicture(F)

GetObject picOrg.Picture, Len(Bb), Bb

If Bb.bmBitsPixel <> 8 Then MsgBox "Bitmap must be 256 colours (8-bit)!": Exit Sub
If Bb.bmWidth <> 200 Or Bb.bmHeight <> 200 Then MsgBox "Bitmap must be 200 pixels by 200 pixels!": Exit Sub

GetBitmapBits picOrg.Picture, UBound(B2), B2(1)

ReDim O(0 To 0) As BitmapOrigin

For X = 0 To 200
For Y = 0 To 200
    If B2((Y * 200) + X) >= 250 Then
    ReDim Preserve O(0 To OCnt) As BitmapOrigin
    O(OCnt).X = X
    O(OCnt).Y = 200 - Y
    OCnt = OCnt + 1
    End If
Next
Next

End Sub

Private Sub Text_Change(Index As Integer)
SetVars
End Sub

Function Shave(S As String) As String
Shave = Mid(S, 1, Len(S) - 4)
End Function

Sub ScanForPals()
Combo2.Clear
Dim D As String
If Dir(App.Path & "\Palettes\", vbDirectory) = "" Then MsgBox "Cannot locate palette directory! No palettes will be avaliable.", vbCritical: Exit Sub
D = Dir(App.Path & "\Palettes\*.Pal")
If D = "" Then MsgBox "No palettes found in palettes directory!", vbCritical: Exit Sub
Do
DoEvents

Combo2.AddItem Shave(D)

D = Dir()
Loop Until D = ""

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Ended = True Then
GoTo Exx
End If
Exit Sub
Exx:
Unload Me
End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Pause"
Command3_Click
Case "About"
Command11_Click
Case "Help"
Command12_Click
Case "Fullscreen"
FrmFS.Move 0, 0, 320 * 15, 200 * 15
FrmFS.Show
ChangeRes 320, 200
FS = True
End Select
End Sub
