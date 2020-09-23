VERSION 5.00
Begin VB.Form lev2 
   BackColor       =   &H0000C000&
   Caption         =   "Racer"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11865
   Icon            =   "lev2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer starttime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4320
      Top             =   6360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   6
      Top             =   7320
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10
      ScaleHeight     =   375
      ScaleWidth      =   2325
      TabIndex        =   5
      Top             =   3360
      Width           =   2325
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      Picture         =   "lev2.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer player2time 
      Enabled         =   0   'False
      Interval        =   9999
      Left            =   3720
      Top             =   6360
   End
   Begin VB.PictureBox Char 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      Picture         =   "lev2.frx":0701
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2400
      Picture         =   "lev2.frx":0B05
      ScaleHeight     =   2295
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image gup 
      Height          =   420
      Left            =   1500
      Picture         =   "lev2.frx":1ED9
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image gdown 
      Height          =   420
      Left            =   1500
      Picture         =   "lev2.frx":22FD
      Top             =   6360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image gleft 
      Height          =   405
      Left            =   1200
      Picture         =   "lev2.frx":272E
      Top             =   6000
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image gright 
      Height          =   390
      Left            =   1800
      Picture         =   "lev2.frx":2B2F
      Top             =   6000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Up 
      Height          =   450
      Left            =   2760
      Picture         =   "lev2.frx":2F33
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Down 
      Height          =   465
      Left            =   2760
      Picture         =   "lev2.frx":3349
      Top             =   6240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Lefta 
      Height          =   390
      Left            =   2400
      Picture         =   "lev2.frx":376E
      Top             =   5760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Right 
      Height          =   450
      Left            =   3120
      Picture         =   "lev2.frx":3B62
      Top             =   5760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Laps left"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   9480
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label timelbl 
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
      Height          =   4095
      Left            =   4200
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Left            =   0
      Top             =   3345
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Laps left"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "3 laps"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image6 
      Height          =   1050
      Left            =   7920
      Picture         =   "lev2.frx":3F59
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   795
      Left            =   4440
      Picture         =   "lev2.frx":46DA
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   5520
      Picture         =   "lev2.frx":4D39
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1050
      Left            =   6720
      Picture         =   "lev2.frx":5393
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   4440
      Picture         =   "lev2.frx":5B14
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   7800
      Picture         =   "lev2.frx":65E6
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   975
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Options"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "lev2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ready As Integer, key As Boolean, mover As Boolean
Dim speed As Integer
Dim lap As Integer
Dim start As Integer
Dim around As Integer
Dim guyl As Integer


Private Sub Command1_Click()
If Command1.Caption = "Start" Then
timelbl.Visible = True
start = 4
starttime.Enabled = True
ElseIf Command1.Caption = "Restart" Then
start = 3
lap = 0
timelbl.Caption = start
Command1.Caption = "Start"
player2time.Enabled = False
Picture2.Top = 840
guyl = 3
Picture2.Picture = LoadPicture("C:\Program Files\DevStudio\VB\fast push\Enl.gif")
Char.Picture = LoadPicture("C:\Program Files\DevStudio\VB\fast push\Guylft.gif")
Picture2.Left = 240
Char.Left = 360
Char.Top = 240
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If key = False And mover = False Then
If KeyCode = vbKeyDown And Char.Top < 7740 Then

If Not (Char.Left < 8760 And Char.Top > 1740) Or Not (Char.Left > 2160 And Char.Top > 1740) Or Char.Top > 4439 Then


Char.Picture = gdown.Picture
Char.Top = Char.Top + speed
End If
End If

If KeyCode = vbKeyUp And Char.Top > 100 Then

If Not (Char.Left < 8760 And Char.Top < 4441) Or Not (Char.Left > 2160 And Char.Top < 4840) Or Char.Top < 2041 Then


If lap = 1 And Char.Top < 3503 And Char.Left < 2161 Then

    If around = 1 Then
     If Mes.hardness = 0 Then
     MsgBox ("You won the woods track" & vbNewLine & "The password is FQDF")
     ElseIf Mes.hardness = 1 Then
     MsgBox ("You won on average" & vbNewLine & "The next password is HYRE")
     ElseIf Mes.hardness = 2 Then
     MsgBox ("Well done, youre getting good." & vbNewLine & "The next password is ATHE")
     ElseIf Mes.hardness = 3 Then MsgBox "Exellent now try the very" & vbNewLine & "way way way too hard mode" & vbNewLine & "with this password F6HY"
     ElseIf Mes.hardness = 4 Then
     MsgBox ("Perfect you won the whole game")
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     Beep
     End
     End If
    Label1.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    player2time.Enabled = False
    Label2.Visible = False
    ElseIf around = 2 Then
    Label2.Caption = "Lap left"
    End If
around = around - 1
Label1.Caption = around
lap = 0
End If

Char.Picture = gup.Picture
Char.Top = Char.Top - speed
End If
End If


If KeyCode = vbKeyRight And Char.Left < 11460 Then

If Char.Left < 2159 Or Char.Top < 2140 Or Char.Top > 4439 Or Char.Left > 8759 Then

Char.Picture = gright.Picture
Char.Left = Char.Left + speed
End If
End If


If KeyCode = vbKeyLeft And Char.Left > 300 Then

If Char.Left < 2161 Or Char.Top > 4439 Or Char.Top < 2041 Or Char.Left > 8761 Then

Char.Picture = gleft.Picture
Char.Left = Char.Left - speed
End If
End If
End If
key = True

If Char.Left > 8759 Then
lap = 1
End If


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Or vbKeyRight Or vbKeyUp Or vbKeyDown Then
key = False
End If
End Sub

Private Sub Form_Load()
If Mes.hardness = 0 Then
around = 2
guyl = 2
player2time.Interval = 300
End If

If Mes.hardness = 1 Then
around = 3
guyl = 3
player2time.Interval = 250
End If

If Mes.hardness = 2 Then
around = 3
guyl = 3
player2time.Interval = 200
End If

If Mes.hardness = 3 Then
around = 4
guyl = 4
player2time.Interval = 160
End If

If Mes.hardness = 4 Then
around = 5
guyl = 5
player2time.Interval = 110
End If


speed = 300
Label1.Caption = around
Label1.Caption = around
Label4.Caption = guyl
key = False
mover = True
End Sub
Private Sub mnuexit_Click()
End
End Sub

Private Sub player2time_Timer()
If Picture2.Left < 8760 And Picture2.Top < 2039 Then
Picture2.Left = Picture2.Left + speed
Picture2.Picture = Right.Picture
DoEvents
ElseIf Picture2.Left > 8759 And Picture2.Top < 4439 Then
Picture2.Top = Picture2.Top + speed
Picture2.Picture = Down.Picture
ElseIf Picture2.Top > 4439 And Picture2.Left > 2159 Then
Picture2.Picture = Lefta.Picture
Picture2.Left = Picture2.Left - speed

ElseIf Picture2.Top > 2039 Then
Picture2.Picture = Up.Picture
Picture2.Top = Picture2.Top - speed
End If

If (Picture2.Top > 2940 And Picture2.Top < 3539) And Picture2.Left < 3841 Then
guyl = guyl - 1
If guyl = 0 Then
player2time.Enabled = False
MsgBox ("Sorry you lost")
Call Command1_Click
ElseIf guyl = 1 Then
Label5.Caption = "Lap left"
End If
Label4.Caption = guyl
End If
End Sub

Private Sub starttime_Timer()
For n = 0 To 1
n = n + 1
DoEvents
start = start - 1
timelbl.Caption = start
If start = 0 Then
starttime.Enabled = False
mover = False
timelbl.Visible = False
Command1.Caption = "Restart"
player2time.Enabled = True
End If
timelbl.Caption = start
Next
End Sub
