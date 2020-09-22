VERSION 5.00
Begin VB.Form fmrMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attack Wave 911    Author: Mr. Brian Rouse        Copyright (c) 2002"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   480
      Width           =   3975
      Begin Brians911.Scroller Scroller1 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1085
         BackColor       =   0
         ForeColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   $"fmrMenu.frx":0000
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Start"
      Height          =   675
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1275
   End
   Begin VB.CheckBox chkSounds 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sounds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   5160
      TabIndex        =   1
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&E&xit"
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Turn Up The Volume!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "fmrMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSounds_Click()

'change sounds
bSounds = Not bSounds

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStart_Click()
'load game
Load frmGame
frmGame.Show
Me.Hide
End Sub

Private Sub Form_Load()
'set defaults
bSounds = True
End Sub

