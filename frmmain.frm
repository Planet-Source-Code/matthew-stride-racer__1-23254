VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00009C31&
   Caption         =   "Racer"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00BDBD06&
      BorderStyle     =   0  'None
      Height          =   7900
      Left            =   10315
      ScaleHeight     =   7905
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   30
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   4080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2.45745e5
      Top             =   2.45745e5
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BDBD06&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   120
      Picture         =   "frmmain.frx":030A
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   360
      Width           =   450
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   120
      Picture         =   "frmmain.frx":06FB
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   3
      Top             =   1200
      Width           =   450
   End
   Begin VB.Image Image33 
      Height          =   1050
      Left            =   8160
      Picture         =   "frmmain.frx":0B44
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   7950
      Left            =   10310
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Hobo BT"
         Size            =   200.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   5400
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image Image32 
      Height          =   1050
      Left            =   6360
      Picture         =   "frmmain.frx":118F
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image31 
      Height          =   1050
      Left            =   4920
      Picture         =   "frmmain.frx":190E
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   855
   End
   Begin VB.Image Image30 
      Height          =   1050
      Left            =   11160
      Picture         =   "frmmain.frx":1F59
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image Image29 
      Height          =   1050
      Left            =   9000
      Picture         =   "frmmain.frx":25BD
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   855
   End
   Begin VB.Image Image28 
      Height          =   1050
      Left            =   5160
      Picture         =   "frmmain.frx":2C21
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image27 
      Height          =   1050
      Left            =   8040
      Picture         =   "frmmain.frx":3285
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   855
   End
   Begin VB.Image Image26 
      Height          =   1050
      Left            =   4320
      Picture         =   "frmmain.frx":38E9
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image Image25 
      Height          =   1050
      Left            =   4320
      Picture         =   "frmmain.frx":3F08
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image Image24 
      Height          =   1050
      Left            =   5400
      Picture         =   "frmmain.frx":456C
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   855
   End
   Begin VB.Image Image23 
      Height          =   255
      Left            =   1080
      Picture         =   "frmmain.frx":4BD0
      Top             =   6240
      Width           =   225
   End
   Begin VB.Image Image22 
      Height          =   255
      Left            =   1440
      Picture         =   "frmmain.frx":4F6A
      Top             =   5280
      Width           =   225
   End
   Begin VB.Image Image21 
      Height          =   255
      Left            =   3960
      Picture         =   "frmmain.frx":5304
      Top             =   6360
      Width           =   225
   End
   Begin VB.Image Image20 
      Height          =   1050
      Left            =   240
      Picture         =   "frmmain.frx":569E
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image19 
      Height          =   1050
      Left            =   1800
      Picture         =   "frmmain.frx":5E1D
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image Image18 
      Height          =   960
      Left            =   2640
      Picture         =   "frmmain.frx":659C
      Top             =   5400
      Width           =   765
   End
   Begin VB.Image Image17 
      Height          =   1050
      Left            =   4200
      Picture         =   "frmmain.frx":6BE7
      Top             =   7080
      Width           =   735
   End
   Begin VB.Image Image16 
      Height          =   1050
      Left            =   7200
      Picture         =   "frmmain.frx":7366
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image Image15 
      Height          =   1050
      Left            =   6000
      Picture         =   "frmmain.frx":7AE5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Image Image13 
      Height          =   720
      Left            =   1080
      Picture         =   "frmmain.frx":8264
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image Image12 
      Height          =   960
      Left            =   2040
      Picture         =   "frmmain.frx":8883
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image Image11 
      Height          =   1050
      Left            =   240
      Picture         =   "frmmain.frx":8ECE
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image Image10 
      Height          =   1050
      Left            =   3480
      Picture         =   "frmmain.frx":964D
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   1050
      Left            =   120
      Picture         =   "frmmain.frx":9DCC
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   1050
      Left            =   960
      Picture         =   "frmmain.frx":A54B
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image Image7 
      Height          =   1050
      Left            =   10680
      Picture         =   "frmmain.frx":ACCA
      Top             =   5880
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   960
      Left            =   3480
      Picture         =   "frmmain.frx":B449
      Top             =   3240
      Width           =   765
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   2640
      Picture         =   "frmmain.frx":BA94
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   1050
      Left            =   3360
      Picture         =   "frmmain.frx":C213
      Top             =   2280
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   1050
      Left            =   7800
      Picture         =   "frmmain.frx":C992
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   1050
      Left            =   6240
      Picture         =   "frmmain.frx":D111
      Top             =   7200
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1800
      Picture         =   "frmmain.frx":D890
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8970
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   855
      Left            =   9345
      Top             =   7020
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   8880
      Top             =   6300
      Width           =   1335
   End
   Begin VB.Image Image14 
      Height          =   1050
      Left            =   7320
      Picture         =   "frmmain.frx":DEDB
      Top             =   4320
      Width           =   735
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Options"
      Begin VB.Menu mnurst 
         Caption         =   "Restart"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnup 
         Caption         =   "Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim press As Boolean
Dim speed As Integer
Dim guy1 As Integer
Dim start As Boolean
Dim startt As Integer
Dim guy1speed As Integer
Public level As Integer
Private Sub Command1_Click()
start = False
If Command1.Caption = "Re&start" Then
Call mnurst_Click
press = False
guy1 = 0
Command1.Caption = "&Start"
ElseIf Command1.Caption = "&Start" Then
Timer2.Enabled = True
Command1.Caption = "Re&start"
Command1.Enabled = False
mnurst.Enabled = False
press = False
mnup.Enabled = False
Command2.Enabled = False
Label1.Visible = True
End If
End Sub

Private Sub Command2_Click()
Call mnup_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If start = True Then
If KeyCode = 39 And press = False And Picture1.Left < 11600 Then
Picture1.Left = Picture1.Left + speed
press = True
End If
End If
If Picture1.Left > Picture2.Left Then
Timer1.Enabled = False
Picture1.Visible = False
Picture3.Visible = False
MsgBox "You won the forest track", vbDefaultButton2
Mes.Label1.Caption = "Keep running in circles"
Mes.Option1.Enabled = False
Mes.Image2.Visible = True
Mes.Image3.Visible = True
Mes.Image4.Visible = True
Mes.Image1.Visible = True
Mes.Option1.Value = False
Mes.Option2.Value = False
Mes.Option3.Value = False
Mes.Option4.Value = False
Mes.Option5.Value = False
Mes.Command4.Enabled = False

If Mes.hardness = 0 Then Mes.Option1.Value = True
If Mes.hardness = 1 Then Mes.Option2.Value = True
If Mes.hardness = 2 Then Mes.Option3.Value = True
If Mes.hardness = 3 Then Mes.Option4.Value = True
If Mes.hardness = 4 Then Mes.Option5.Value = True
If Mes.hardness = 5 Then Mes.Option5.Value = True
Mes.Show
Unload frmmain
Mes.Caption = "Woods Track"
Mes.Label1.Caption = "Do five laps around the house and it."
Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
press = False
End Sub

Private Sub Form_Load()
level = 1
press = False
speed = 250
start = False
Label1.Visible = False
guy1speed = 95
Command2.Enabled = False
If Mes.hardness = 0 Then
Timer1.Interval = 300
ElseIf Mes.hardness = 1 Then
Timer1.Interval = 200
ElseIf Mes.hardness = 2 Then
Timer1.Interval = 150
ElseIf Mes.hardness = 3 Then
Timer1.Interval = 145
ElseIf Mes.hardness = 4 Then
Timer1.Interval = 100
End If
End Sub
Private Sub mnuav_Click()
hard = 2
mnueasy.Checked = False
guy1speed = 95
mnuav.Checked = True
mnuhard.Checked = False
mnuwth.Checked = False
End Sub

Private Sub mnucity_Click(Index As Integer)
lev2.Show
End Sub

Private Sub mnueasy_Click()
hard = 1
mnueasy.Checked = True
guy1speed = 75
mnuav.Checked = False
mnuhard.Checked = False
mnuwth.Checked = False
End Sub



Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuhard_Click()
hard = 3
mnueasy.Checked = False
guy1speed = 120
mnuav.Checked = False
mnuhard.Checked = True
mnuwth.Checked = False
End Sub

Private Sub mnup_Click()
If mnup.Caption = "Pause" Then
Command2.Caption = "&Resume"
mnup.Caption = "Resume"
Timer1.Enabled = False
start = False
Else
Command2.Caption = "&Pause"
mnup.Caption = "Pause"
Timer1.Enabled = True
start = True
End If
End Sub

Private Sub mnurst_Click()
Picture1.Top = 360
Picture1.Visible = True
Picture3.Visible = True
Label1.Caption = 3
Picture1.Left = 120
Picture3.Left = 120
Label1.Visible = False
press = False
guy1 = 0
speed = 300
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub mnuwth_Click()
hard = 4
mnueasy.Checked = False
guy1speed = 150
Timer1.Interval = 20
mnuav.Checked = False
mnuhard.Checked = False
mnuwth.Checked = True
End Sub



Private Sub Timer1_Timer()
If Picture3.Left > Picture2.Left Then
Timer1.Enabled = False
Picture1.Visible = False
Picture3.Visible = False
MsgBox "You loose", vbCritical
Exit Sub
End If
guy1 = guy1 + guy1speed
Picture3.Left = Picture3.Left + 300
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = 0 Then
Timer1.Enabled = True
press = True
Command2.Enabled = True
Timer2.Enabled = False
Command1.Enabled = True
mnurst.Enabled = True
mnup.Enabled = True
Command2.Enabled = True
Command1.Caption = "Re&start"
start = True
Label1.Visible = False
Else
Label1.Caption = Label1.Caption - 1
End If
End Sub
