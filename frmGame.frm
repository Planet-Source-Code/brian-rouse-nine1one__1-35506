VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Brian's Attack Wave 911"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   ControlBox      =   0   'False
   FillStyle       =   7  'Diagonal Cross
   Icon            =   "frmGame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmGame.frx":0442
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrStatWait 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3900
      Top             =   4110
   End
   Begin VB.Timer tmrWaveTime 
      Interval        =   40000
      Left            =   4590
      Top             =   4110
   End
   Begin VB.Timer tmrNuke 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   300
      Top             =   1080
   End
   Begin VB.Timer tmrICBMAnim 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   25
      Left            =   6690
      Top             =   480
   End
   Begin VB.Timer tmrMissile3Anim 
      Interval        =   60
      Left            =   1440
      Top             =   1800
   End
   Begin VB.Timer tmrMissile2Anim 
      Interval        =   60
      Left            =   1440
      Top             =   1140
   End
   Begin VB.Timer tmrMissile1Anim 
      Interval        =   60
      Left            =   1410
      Top             =   510
   End
   Begin VB.Label lblComplete 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Next Bin-Laden Attack"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1005
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   8220
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mr. Brian A. Rouse, Author/Developer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   870
      TabIndex        =   4
      Top             =   3600
      Width           =   3705
   End
   Begin VB.Label lblScorelabel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   5310
      TabIndex        =   3
      Top             =   4680
      Width           =   630
   End
   Begin VB.Line linGround 
      BorderColor     =   &H00FFFFFF&
      X1              =   540
      X2              =   7980
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label lblICBMsLeft 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6840
      TabIndex        =   2
      Top             =   4260
      Width           =   75
   End
   Begin VB.Label lblICBMInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nuclear War Heads Remaining"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   4680
      TabIndex        =   1
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Shape recTarget 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      Height          =   1455
      Index           =   0
      Left            =   4410
      Top             =   1050
      Width           =   465
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6840
      TabIndex        =   0
      Top             =   4650
      Width           =   1305
   End
   Begin VB.Shape cirICBM 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   165
      Index           =   0
      Left            =   6750
      Shape           =   2  'Oval
      Top             =   990
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Line linMissile3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   0
      X1              =   2160
      X2              =   2400
      Y1              =   1860
      Y2              =   2100
   End
   Begin VB.Line linMissile1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2010
      X2              =   2310
      Y1              =   570
      Y2              =   900
   End
   Begin VB.Line linMissile2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   0
      X1              =   2070
      X2              =   2340
      Y1              =   1170
      Y2              =   1410
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim iNumMissiles As Integer
Dim iNumICBMs As Integer
Dim bShrink(1 To 5) As Integer
Dim bTargetStatus(0 To 9) As Boolean
Dim bPosInUse(1 To 5) As Boolean
Dim iLastSeg(1 To 3) As Integer
Dim iLeftRnd(1 To 3), iRightRnd(1 To 3) As Integer
Dim iR1, iL1, iR2, iL2, iR3, iL3 As Integer



Private Sub Form_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 27 'escape (and exit)
        End
End Select

End Sub

Private Sub Form_Load()
Dim rc As Long
Dim iCount As Integer
Randomize

'set screen dimentions
With Me
    .Width = Screen.Width
    .Height = Screen.Height
    .Top = 0
    .Left = 0
End With
With linGround
    .Y1 = Me.Height - 1000
    .Y2 = Me.Height - 1000
    .X1 = 0
    .X2 = Me.Width
End With
With lblScore
    .Left = Me.Width - (.Width + 100)
    .Top = Me.Height - (.Height + 100)
End With
With lblICBMInfo
    .Left = Me.Width - (.Width + 2000)
    .Top = Me.Height - (.Height + 600)
End With
With lblICBMsLeft
    .Left = Me.Width - (.Width + 1000)
    .Top = Me.Height - (.Height + 600)
End With
With lblScorelabel
    .Left = lblICBMInfo.Left
    .Top = Me.Height - (.Height + 200)
End With
With lblScore
    .Left = lblICBMsLeft.Left
    .Top = Me.Height - (.Height + 200)
End With
With lblTitle
    .Left = 100
    .Top = Me.Height - (.Height + 200)
End With
With lblComplete
    .Top = (Me.Height / 2) - (.Height / 2)
    .Left = (Me.Width / 2) - (.Width / 2)
End With
    

'initialise level 1 values
iVertSpeed = 200
iICBMDiam = 1000
iLevel = 1
bICBMOK = True
bNukeOK = True
bGameOn = True
iNumICBMs = 0
iEMPsFired = 0
iEMPsAllowed = 0
bHoming = False
iNumMissiles = 3
iScore = 0
iNumHit = 0
iNumFired = 0
iNumAllowed = 30
LoadTargets 5

lblICBMsLeft.Caption = iNumAllowed

'initialise missiles and targets

For iCount = 1 To 3
    iRightRnd(iCount) = 100
    iLeftRnd(iCount) = 100
Next iCount
For iCount = 0 To iNumTargets - 1
    bTargetStatus(iCount) = True
Next iCount

Init (0)
Init (1)
Init (2)

If bSounds Then rc = sndPlaySound(App.Path & "\subdive.wav", 1)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then  'launch ABM
Dim iNewPos As Integer
Dim iFreePos As Integer
Dim sWavFile As String
Dim iSnd As Integer
Dim rc As Long
'create new ABM
If (iNumICBMs <= 4) And (iNumFired < iNumAllowed) And bICBMOK And (Y < (linGround.Y1 - iICBMDiam)) Then
    
    For iNewPos = 1 To 5
        If bPosInUse(iNewPos) = False Then
            iFreePos = iNewPos
            Exit For
        End If
    Next iNewPos
    
    bPosInUse(iFreePos) = True
    iNumFired = iNumFired + 1
    lblICBMsLeft.Caption = iNumAllowed - iNumFired
    
    'load new circle
    iNumICBMs = iNumICBMs + 1
    Load cirICBM(iFreePos)
    
    'initialise circle
    With cirICBM(iFreePos)
        .Left = X - (.Width / 2)
        .Top = Y - (.Height / 2)
        .Visible = True
    End With
    'load new animation timer for the circle
    Load tmrICBMAnim(iFreePos)
    'initialise timer
    With tmrICBMAnim(iFreePos)
        .Interval = 25
        .Enabled = True
    End With
    'play a random blast sound
    
    If bSounds Then
    iSnd = Fix(Rnd * (4 - 1) + 1)
    Select Case iSnd
        Case 1
            sWavFile = App.Path & "\boom.wav"
        Case 2
            sWavFile = App.Path & "\explode.wav"
        Case 3
            sWavFile = App.Path & "\explosion.wav"
        Case 4
            sWavFile = App.Path & "\explosion2.wav"
    End Select
    'play sound
    rc = sndPlaySound(sWavFile, 1)
    End If
        
    
End If
Else
  If bNukeOK And (iEMPsFired < iEMPsAllowed) Then
    'launch NUKE!
    iEMPsFired = iEMPsFired + 1
    'clear all missiles and flash screen
    tmrMissile1Anim.Enabled = False
    tmrMissile2Anim.Enabled = False
    tmrMissile3Anim.Enabled = False
    
    Dim iSeg As Integer
    For iSeg = 1 To iLastSeg(1)
        Unload linMissile1(iSeg)
    Next iSeg
    iLastSeg(1) = 0
    Init (0)
    For iSeg = 1 To iLastSeg(2)
        Unload linMissile2(iSeg)
    Next iSeg
    iLastSeg(2) = 0
    Init (1)
    For iSeg = 1 To iLastSeg(3)
        Unload linMissile3(iSeg)
    Next iSeg
    iLastSeg(3) = 0
    Init (2)
    Me.BackColor = RGB(255, 255, 50)
    lblScore.BackColor = RGB(255, 255, 50)
    lblICBMsLeft.BackColor = RGB(255, 255, 50)
    lblScorelabel.BackColor = RGB(255, 255, 50)
    lblICBMInfo.BackColor = RGB(255, 255, 50)
    tmrNuke.Enabled = True
    
    If bSounds Then rc = sndPlaySound(App.Path & "\dissolve2.wav", 1)
  End If
End If

End Sub

Private Sub tmrICBMAnim_Timer(Index As Integer)

'check to see if it strikes a target
Check4Hit Index

If Not bShrink(Index) Then
    'expand icbm
    With cirICBM(Index)
        .Height = .Height + 30
        .Width = .Width + 30
        .Left = .Left - 15
        .Top = .Top - 15
        
        If .Width > iICBMDiam Then bShrink(Index) = True
    End With
Else
    'shrink icbm
    With cirICBM(Index)
        .Height = .Height - 30
        .Width = .Width - 30
        .Left = .Left + 15
        .Top = .Top + 15
    End With
        'remove icbm once shrunk
        If cirICBM(Index).Width < 165 Then
            Unload cirICBM(Index)
            bPosInUse(Index) = False
            bShrink(Index) = False
            iNumICBMs = iNumICBMs - 1
            Unload tmrICBMAnim(Index)
        End If

End If

End Sub

Private Sub tmrMissile1Anim_Timer()


Static bRight As Boolean

Dim iSeg As Integer

'increment number of trace segments
iLastSeg(1) = iLastSeg(1) + 1
'alternate left/right
bRight = Not bRight

'create new segment
Load linMissile1(iLastSeg(1))
With linMissile1(iLastSeg(1))
    'adjoin to previous and randomise end pos
    .X1 = linMissile1(iLastSeg(1) - 1).X2
    .Y1 = linMissile1(iLastSeg(1) - 1).Y2
    
    If bRight Then
        .X2 = linMissile1(iLastSeg(1) - 1).X2 + Rnd * (iRightRnd(1) - 10) + 10
    Else
        .X2 = linMissile1(iLastSeg(1) - 1).X2 - Rnd * (iLeftRnd(1) - 10) + 10
    End If
    
    .Y2 = linMissile1(iLastSeg(1) - 1).Y2 + Rnd * (iVertSpeed - 10) + 10
    .Visible = True

    'fade red component of previous segments
    For iSeg = 1 To iLastSeg(1)
        If linMissile1(iSeg).BorderColor > RGB(0, 0, 0) Then
            linMissile1(iSeg).BorderColor = linMissile1(iSeg).BorderColor - RGB(5, 0, 0)
        End If
    Next iSeg
    
    'check if missile has hit bottom
    If .Y2 >= linGround.Y1 Then
        For iSeg = 1 To iLastSeg(1)
        'unload missile t
            Unload linMissile1(iSeg)
            iLastSeg(1) = 0
        Next iSeg
        'initialise new missile pos
        Init (0)
    End If
    
    'check if missile has collided with a target
    If linMissile1(iLastSeg(1)).Y2 >= recTarget(0).Top Then
        CheckTargets 1
    End If
End With


End Sub

Private Sub Init(iMisNum)
'initialise first missile segment of each missile and pick target
Dim iTagNum, iCount, iBaseToTarget As Integer
Dim iTagArr(0 To 9) As Integer
iCount = 0
For iTagNum = 1 To iNumTargets
    If bTargetStatus(iTagNum - 1) = True Then
        iTagArr(iCount) = iTagNum
        iCount = iCount + 1
    End If
Next iTagNum

iBaseToTarget = iTagArr(Fix(Rnd * iCount))

'MsgBox "Missile " & iMisNum + 1 & " has targeted base " & iBaseToTarget
'MsgBox bTargetStatus(0) & bTargetStatus(1) & bTargetStatus(2) & bTargetStatus(3) & bTargetStatus(4)

Select Case iMisNum
    Case 0
        With linMissile1(0)
            .X1 = Rnd * (Me.Width - 0) + 0
            .Y1 = 0
            .Y2 = 0
            .X2 = .X1
        End With
        iLastSeg(1) = 0
        If bHoming Then
            HomeOnTarget 1, iBaseToTarget
        End If
    Case 1
        With linMissile2(0)
            .X1 = Rnd * (Me.Width - 0) + 0
            .Y1 = 0
            .Y2 = 0
            .X2 = .X1
        End With
        iLastSeg(2) = 0
        If bHoming Then
            HomeOnTarget 2, iBaseToTarget
        End If
    Case 2
        With linMissile3(0)
            .X1 = Rnd * (Me.Width - 0) + 0
            .Y1 = 0
            .Y2 = 0
            .X2 = .X1
        End With
        iLastSeg(3) = 0
        If bHoming Then
            HomeOnTarget 3, iBaseToTarget
        End If
End Select
End Sub

Private Sub tmrMissile2Anim_Timer()

Static bRight As Boolean

Dim iSeg As Integer

'increment number of trace segments
iLastSeg(2) = iLastSeg(2) + 1
'alternate left/right
bRight = Not bRight

'create new segment
Load linMissile2(iLastSeg(2))
With linMissile2(iLastSeg(2))
    'adjoin to previous and randomise end pos
    .X1 = linMissile2(iLastSeg(2) - 1).X2
    .Y1 = linMissile2(iLastSeg(2) - 1).Y2
    
    If bRight Then
        .X2 = linMissile2(iLastSeg(2) - 1).X2 + Rnd * (iRightRnd(2) - 10) + 10
    Else
        .X2 = linMissile2(iLastSeg(2) - 1).X2 - Rnd * (iLeftRnd(2) - 10) + 10
    End If
    
    .Y2 = linMissile2(iLastSeg(2) - 1).Y2 + Rnd * (iVertSpeed - 10) + 10
    .Visible = True

    'fade red component of previous segments
    For iSeg = 1 To iLastSeg(2)
        If linMissile2(iSeg).BorderColor > RGB(0, 0, 0) Then
            linMissile2(iSeg).BorderColor = linMissile2(iSeg).BorderColor - RGB(0, 0, 5)
        End If
    Next iSeg
    
    'check if missile has hit bottom
    If .Y2 >= linGround.Y1 Then
        For iSeg = 1 To iLastSeg(2)
        'unload missile t
            Unload linMissile2(iSeg)
            iLastSeg(2) = 0
        Next iSeg
        'initialise new missile pos
        Init (1)
    End If
    
    'check if missile has collided with a target
    If linMissile2(iLastSeg(2)).Y2 >= recTarget(0).Top Then
        CheckTargets 2
    End If
    
End With


End Sub

Private Sub tmrMissile3Anim_Timer()

Static bRight As Boolean

Dim iSeg As Integer

'increment number of trace segments
iLastSeg(3) = iLastSeg(3) + 1
'alternate left/right
bRight = Not bRight

'create new segment
Load linMissile3(iLastSeg(3))
With linMissile3(iLastSeg(3))
    'adjoin to previous and randomise end pos
    .X1 = linMissile3(iLastSeg(3) - 1).X2
    .Y1 = linMissile3(iLastSeg(3) - 1).Y2
    
    If bRight Then
        .X2 = linMissile3(iLastSeg(3) - 1).X2 + Rnd * (iRightRnd(3) - 10) + 10
    Else
        .X2 = linMissile3(iLastSeg(3) - 1).X2 - Rnd * (iLeftRnd(3) - 10) + 10
    End If
    
    .Y2 = linMissile3(iLastSeg(3) - 1).Y2 + Rnd * (iVertSpeed - 10) + 10
    .Visible = True

    'fade red component of previous segments
    For iSeg = 1 To iLastSeg(3)
        If linMissile3(iSeg).BorderColor > RGB(0, 0, 0) Then
            linMissile3(iSeg).BorderColor = linMissile3(iSeg).BorderColor - RGB(0, 5, 0)
        End If
    Next iSeg
    
    'check if missile has hit bottom
    If .Y2 >= linGround.Y1 Then
        For iSeg = 1 To iLastSeg(3)
        'unload missile t
            Unload linMissile3(iSeg)
            iLastSeg(3) = 0
        Next iSeg
        'initialise new missile pos
        Init (2)
    End If
    
    'check if missile has collided with a target
    If linMissile3(iLastSeg(3)).Y2 >= recTarget(0).Top Then
        CheckTargets 3
    End If
End With

End Sub

Private Sub Check4Hit(ICBMIndex As Integer)
Dim iSeg As Integer
Dim bHit As Boolean
bHit = False

'check against red missile
With linMissile1(iLastSeg(1))
    If (.X1 >= cirICBM(ICBMIndex).Left) And (.X2 <= (cirICBM(ICBMIndex).Left + cirICBM(ICBMIndex).Width)) Then
        If (.Y2 >= cirICBM(ICBMIndex).Top) And (.Y2 <= (cirICBM(ICBMIndex).Top + cirICBM(ICBMIndex).Height)) Then
            'ICBM has hit missile 1 (red)
            bHit = True
            For iSeg = 1 To iLastSeg(1)
                Unload linMissile1(iSeg)
                iLastSeg(1) = 0
            Next
            Init (0)
        End If
    End If
End With
'check against blue missile
With linMissile2(iLastSeg(2))
    If (.X1 >= cirICBM(ICBMIndex).Left) And (.X2 <= (cirICBM(ICBMIndex).Left + cirICBM(ICBMIndex).Width)) Then
        If (.Y2 >= cirICBM(ICBMIndex).Top) And (.Y2 <= (cirICBM(ICBMIndex).Top + cirICBM(ICBMIndex).Height)) Then
            'ICBM has hit missile 2 (blue)
            bHit = True
            For iSeg = 1 To iLastSeg(2)
                Unload linMissile2(iSeg)
                iLastSeg(2) = 0
            Next
            Init (1)
        End If
    End If
End With
'check against green missile
With linMissile3(iLastSeg(3))
    If (.X1 >= cirICBM(ICBMIndex).Left) And (.X2 <= (cirICBM(ICBMIndex).Left + cirICBM(ICBMIndex).Width)) Then
        If (.Y2 >= cirICBM(ICBMIndex).Top) And (.Y2 <= (cirICBM(ICBMIndex).Top + cirICBM(ICBMIndex).Height)) Then
            'ICBM has hit missile 3 (green)
            bHit = True
            For iSeg = 1 To iLastSeg(3)
                Unload linMissile3(iSeg)
                iLastSeg(3) = 0
            Next
            Init (2)
        End If
    End If
End With

If bHit Then
    iScore = iScore + (iLevel * 100)
    lblScore.Caption = iScore
    iNumSucceed = iNumSucceed + 1
End If
 
    
End Sub



Private Sub LoadTargets(iNum2Make)
Dim iCount, iSep As Integer

iSep = 0
recTarget(0).Left = (Me.Width / (2 * iNum2Make)) - (recTarget(0).Width / 2)
recTarget(0).Top = linGround.Y1 - recTarget(0).Height
If iNum2Make > 1 Then
For iCount = 2 To iNum2Make
    Load recTarget(iCount - 1)
    With recTarget(iCount - 1)
        .Top = linGround.Y1 - .Height
        .Left = recTarget(iCount - 2).Left + (2 * (Me.Width / (2 * iNum2Make)))
        .Visible = True
    End With
Next iCount
End If
iNumTargets = iNum2Make
    
End Sub

Private Sub CheckTargets(iMisNum As Integer)
Dim iCount, iSeg As Integer
Dim bHit As Boolean
Dim rc As Long

Select Case iMisNum
    Case 1
        For iCount = 0 To (iNumTargets - 1)
            With linMissile1(iLastSeg(iMisNum))
            If (.X2 >= recTarget(iCount).Left) And (.X2 <= (recTarget(iCount).Left + recTarget(iCount).Width)) And recTarget(iCount).Visible = True Then
                recTarget(iCount).Visible = False
                bTargetStatus(iCount) = False
                bHit = True
                For iSeg = 1 To iLastSeg(iMisNum)
                    Unload linMissile1(iSeg)
                Next iSeg
                iLastSeg(1) = 0
                Init (0)
                Exit For
            End If
            End With
        Next iCount
        
    Case 2
         For iCount = 0 To (iNumTargets - 1)
            With linMissile2(iLastSeg(iMisNum))
            If (.X2 >= recTarget(iCount).Left) And (.X2 <= (recTarget(iCount).Left + recTarget(iCount).Width)) And recTarget(iCount).Visible = True Then
                recTarget(iCount).Visible = False
                bTargetStatus(iCount) = False
                bHit = True
                For iSeg = 1 To iLastSeg(iMisNum)
                    Unload linMissile2(iSeg)
                Next iSeg
                iLastSeg(2) = 0
                Init (1)
                Exit For
            End If
            End With
        Next iCount
        
    Case 3
         For iCount = 0 To (iNumTargets - 1)
            With linMissile3(iLastSeg(iMisNum))
            If (.X2 >= recTarget(iCount).Left) And (.X2 <= (recTarget(iCount).Left + recTarget(iCount).Width)) And recTarget(iCount).Visible = True Then
                bHit = True
                recTarget(iCount).Visible = False
                bTargetStatus(iCount) = False
                For iSeg = 1 To iLastSeg(iMisNum)
                    Unload linMissile3(iSeg)
                Next iSeg
                iLastSeg(3) = 0
                Init (2)
                Exit For
            End If
            End With
        Next iCount
End Select

'missile has destroyed building
If bHit Then
    iNumHit = iNumHit + 1
    'play alarm
    If bSounds Then rc = sndPlaySound(App.Path & "\alarm02.wav", 1)
End If

If iNumHit >= iNumTargets Then
    GameOver
End If

End Sub

Private Sub GameOver()
    MsgBox "Congratulations, you have just destroyed humanity!"
End Sub

Private Sub tmrNuke_Timer()
'fade screen to black
If (Me.BackColor - RGB(15, 15, 0)) > RGB(0, 0, 50) Then
    Me.BackColor = Me.BackColor - RGB(15, 15, 0)
    lblICBMsLeft.BackColor = lblICBMsLeft.BackColor - RGB(15, 15, 0)
    lblICBMInfo.BackColor = lblICBMInfo.BackColor - RGB(15, 15, 0)
    lblScorelabel.BackColor = lblScorelabel.BackColor - RGB(15, 15, 0)
    lblScore.BackColor = lblScore.BackColor - RGB(15, 15, 0)
Else
    Me.BackColor = RGB(0, 0, 0)
    lblScore.BackColor = RGB(0, 0, 0)
    lblICBMsLeft.BackColor = RGB(0, 0, 0)
    lblScorelabel.BackColor = RGB(0, 0, 0)
    lblICBMInfo.BackColor = RGB(0, 0, 0)
    If bGameOn Then
        tmrMissile1Anim.Enabled = True
        tmrMissile2Anim.Enabled = True
        tmrMissile3Anim.Enabled = True
    End If
    tmrNuke.Enabled = False
End If
End Sub

Private Sub tmrStatWait_Timer()

lblComplete.Visible = False
'show wave statistics
Load frmWaveStats
frmWaveStats.Show 1
tmrStatWait.Enabled = False

End Sub

Private Sub tmrWaveTime_Timer()
Dim rc As Long
Dim iMis As Integer

'disable firing
bICBMOK = False
bNukeOK = False
'play sound
If bSounds Then rc = sndPlaySound(App.Path & "\buzzer.wav", 1)
tmrMissile1Anim.Enabled = False
tmrMissile2Anim.Enabled = False
tmrMissile3Anim.Enabled = False
For iMis = 1 To iNumMissiles
    UnloadMissile (iMis)
Next iMis

lblComplete.Visible = True
tmrStatWait.Enabled = True
tmrWaveTime.Enabled = False
End Sub
Private Sub UnloadMissile(iMisNum)
Dim iSeg As Integer

Select Case iMisNum
    Case 1
        For iSeg = 1 To iLastSeg(1)
            Unload linMissile1(iSeg)
        Next iSeg
        iLastSeg(1) = 0
        Init (0)
    Case 2
        For iSeg = 1 To iLastSeg(2)
            Unload linMissile2(iSeg)
        Next iSeg
        iLastSeg(2) = 0
        Init (1)
    Case 3
        For iSeg = 1 To iLastSeg(3)
            Unload linMissile3(iSeg)
        Next iSeg
        iLastSeg(3) = 0
        Init (2)
End Select

End Sub

Private Sub HomeOnTarget(iMisNum As Integer, iTarget As Integer)

'adjust for arraypos
iTarget = iTarget - 1

Select Case iMisNum
    Case 1
        If linMissile1(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) > 0 Then
            'target is to the left of missile launch
            iRightRnd(iMisNum) = 10
            iLeftRnd(iMisNum) = Fix(2.2 * (iVertSpeed / (Me.Height / (linMissile1(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2))))))
        ElseIf linMissile1(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) < 0 Then
            'target is to the right of missile launch
            iRightRnd(iMisNum) = Fix(2.2 * (iVertSpeed / (Me.Height / ((recTarget(iTarget).Left + (recTarget(iTarget).Width / 2) - linMissile1(0).X1)))))
            iLeftRnd(iMisNum) = 10
        Else
            'target is directly underneath missile launch (rare!)
        End If
    
    Case 2
        If linMissile2(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) > 0 Then
            'target is to the left of missile launch
            iRightRnd(iMisNum) = 10
            iLeftRnd(iMisNum) = Fix(2.2 * (iVertSpeed / (Me.Height / (linMissile2(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2))))))
        ElseIf linMissile2(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) < 0 Then
            'target is to the right of missile launch
            iRightRnd(iMisNum) = 2.2 * (iVertSpeed / (Me.Height / ((recTarget(iTarget).Left + (recTarget(iTarget).Width / 2) - linMissile2(0).X1))))
            iLeftRnd(iMisNum) = 10
        Else
            'target is directly underneath missile launch (rare!)
        End If
    
    Case 3
        If linMissile3(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) > 0 Then
            'target is to the left of missile launch
            iRightRnd(iMisNum) = 10
            iLeftRnd(iMisNum) = Fix(2.2 * (iVertSpeed / (Me.Height / (linMissile3(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2))))))
        ElseIf linMissile3(0).X1 - (recTarget(iTarget).Left + (recTarget(iTarget).Width / 2)) < 0 Then
            'target is to the right of missile launch
            iRightRnd(iMisNum) = 2.2 * (iVertSpeed / (Me.Height / ((recTarget(iTarget).Left + (recTarget(iTarget).Width / 2) - linMissile3(0).X1))))
            iLeftRnd(iMisNum) = 10
        Else
            'target is directly underneath missile launch (rare!)
        End If
    
    
End Select

End Sub
