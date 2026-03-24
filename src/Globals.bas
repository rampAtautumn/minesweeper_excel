Option Explicit

'=======================
' GAME STATE VARIABLES
'=======================

Public GameRunning As Boolean
Public GamePaused As Boolean
Public GameEnded As Boolean
Public CurrentRound As Integer
Public Const MaxRound As Integer = 20

Public ReloadTime As Double
Public LastShotTime As Double
Public DeltaTime As Double
Public LastFrameTime As Double

'=======================
' SCORE / PROGRESS
'=======================

Public Score As Long
Public DucksShot As Integer
Public DucksMissed As Integer

'=======================
' PLAYER / WEAPON
'=======================

Public Bullets As Integer
Public Const MaxBullets As Integer = 3
Public PlayerShot As Boolean

'=======================
' MOUSE POSITION
'=======================

Public MouseX As Double
Public MouseY As Double

'=======================
' DUCKS
'=======================

Public Ducks As Collection
Public DucksPerRound As Integer
Public DucksSpawned As Integer
Public SpawnDelay As Double
Public LastSpawnTime As Double

'=======================
' GAME SPEED AND FRAME
'=======================

Public GameSpeed As Double
Public Const FrameDelay As Double = 0.0333

'=======================
' SHEET REFERENCES
'=======================

Public MenuSheet As Worksheet
Public GameSheet As Worksheet
Public PauseSheet As Worksheet
Public SpriteSheet As Worksheet

'=======================
' SHEET NAMES (CONSTANTS)
'=======================

Public Const SHEET_GAME As String = "Game"
Public Const SHEET_MENU As String = "Menu"
Public Const SHEET_SPRITES As String = "Sprites"
Public Const SHEET_PAUSE As String = "Pause"
