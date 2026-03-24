Option Explicit

'=======================
' GAME STATE VARIABLES
'=======================

Public GameRunning As Boolean
Public GamePaused As Boolean
Public GameEnded As Boolean
Public CurrentRound As Integer
Public Const MaxRound As Integer = 20

'=======================
' TIMING
'=======================

Public DeltaTime As Double
Public LastFrameTime As Double
Public Const FrameDelay As Double = 0.0333 ' ~30 FPS

Public ReloadTime As Double
Public LastShotTime As Double

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
' GAME SPEED
'=======================

Public GameSpeed As Double

'=======================
' SHEET REFERENCES
'=======================

Public GameSheet As Worksheet
Public MenuSheet As Worksheet
Public PauseSheet As Worksheet

'=======================
' SHEET NAMES (CONSTANTS)
'=======================

Public Const SHEET_GAME As String = "Game"
Public Const SHEET_MENU As String = "Menu"
Public Const SHEET_PAUSE As String = "Pause"
Public Const SHEET_SPRITES As String = "GameScreen"

'=======================
' INITIALIZATION
'=======================

Public Sub InitializeGlobals()
    ' Initialize worksheet references
    On Error Resume Next
    Set GameSheet = ThisWorkbook.Sheets(SHEET_GAME)
    Set MenuSheet = ThisWorkbook.Sheets(SHEET_MENU)
    Set PauseSheet = ThisWorkbook.Sheets(SHEET_PAUSE)
    On Error GoTo 0
    
    ' Initialize game variables
    GameRunning = False
    GamePaused = False
    GameEnded = False
    CurrentRound = 1
    Score = 0
    DucksShot = 0
    DucksMissed = 0
    Bullets = MaxBullets
    PlayerShot = False
    DucksPerRound = 5
    DucksSpawned = 0
    GameSpeed = 1.0
    MouseX = 0
    MouseY = 0
End Sub