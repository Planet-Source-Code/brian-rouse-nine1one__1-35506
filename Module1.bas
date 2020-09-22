Attribute VB_Name = "Module1"

'global declarations

'Sound API Declaration
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'sounds on/off
Public bSounds As Boolean

'number of ground targets currently set
Public iNumTargets As Integer

'number of ground targets destroyed
Public iNumHit As Integer

'number of ABMs fired in level
Public iNumFired As Integer

'players total score
Public iScore As Integer

'number of EMPs fired in level
Public iEMPsFired As Integer

'current level of play (attack wave)
Public iLevel As Integer

'apx. vertical speed of missiles
Public iVertSpeed As Integer

'number of successful ABM strikes
Public iNumSucceed As Integer

'number of ABMs available in level
Public iNumAllowed As Integer

'number of EMPs available in level
Public iEMPsAllowed As Integer

'blast diameter of ABM/AAG
Public iICBMDiam As Integer

'if the missiles home on a target or not
Public bHoming As Boolean

'if shooting is allowed
Public bGameOn As Boolean
Public bICBMOK, bNukeOK As Boolean

