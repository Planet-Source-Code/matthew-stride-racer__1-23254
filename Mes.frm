VERSION 5.00
Begin VB.Form Mes 
   BackColor       =   &H00C0C000&
   Caption         =   "Forest Track"
   ClientHeight    =   5400
   ClientLeft      =   2550
   ClientTop       =   2145
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7140
   Begin VB.CommandButton Command4 
      Caption         =   "Password"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Directions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3600
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1440
         Picture         =   "Mes.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1440
         Picture         =   "Mes.frx":030A
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   960
         Picture         =   "Mes.frx":0614
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1920
         Picture         =   "Mes.frx":091E
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF00&
         Caption         =   "Way, way way to hard"
         Enabled         =   0   'False
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF00&
         Caption         =   "Way too hard"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Hard"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Average"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Easy"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Objectives"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   240
         Shape           =   3  'Circle
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Make it to the finish line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "Mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public hardness As Integer
Private Sub Command1_Click()
If Mes.Caption = "Forest Track" Then
frmmain.Show
ElseIf Mes.Caption = "Woods Track" Then
lev2.Show
End If
Unload Mes
End Sub

Private Sub Command2_Click()
Helpfrm.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
resp = InputBox("Enter Password")
resp = UCase(resp)
If resp = "FQDF" Then
Option2.Enabled = True
ElseIf resp = "HYRE" Then
Option2.Enabled = True
Option3.Enabled = True
ElseIf resp = "AHTE" Then
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
ElseIf resp = "F6HY" Then
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Else: MsgBox "Invalid Password"
End If
End Sub


Private Sub Option1_Click()
Mes.hardness = 0
End Sub

Private Sub Option2_Click()
Mes.hardness = 1
End Sub

Private Sub Option3_Click()
Mes.hardness = 2
End Sub

Private Sub Option4_Click()
Mes.hardness = 3
End Sub

Private Sub Option5_Click()
Mes.hardness = 4
End Sub
