Option Explicit

'=======================
' GAME STATE VARIABLES
'=======================

Public GameRunning As Boolean
Public GamePaused As Boolean
Public GameEnded As Boolean
Public CurrentRound As Long
Public Const MaxRound As Long = 20

'=======================
' SPRITES PATHS
'=======================

Public Const ASSETS_ROOT As String = "Assets\sprites\"

Public Const PATH_DUCKS As String = "Sprites patos\"
Public Const PATH_DOG As String = "Sprites perro\"
Public Const PATH_BACKGROUNDS As String = "Fondos y otros\"

'=======================
' TIMING
'=======================

Public DeltaTime As Double
Public LastFrameTime As Double
Public Const FrameDelay As Double = 0.0333

Public ReloadTime As Double
Public LastShotTime As Double
Public LastSpawnTime As Double

'=======================
' SCORE / PROGRESS
'=======================

Public Score As Long
Public DucksShot As Long
Public DucksMissed As Long

'=======================
' PLAYER / WEAPON
'=======================

Public Bullets As Long
Public Const MaxBullets As Long = 3
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
Public DucksPerRound As Long
Public DucksSpawned As Long
Public SpawnDelay As Double

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
' SHEET NAMES
'=======================

Public Const SHEET_GAME As String = "Game"
Public Const SHEET_MENU As String = "Menu"
Public Const SHEET_PAUSE As String = "Pause"
Public Const SHEET_SPRITES As String = "GameScreen"

'=======================
' INITIALIZATION
'=======================

Public Sub InitializeGlobals()

    ' Sheets
    Set GameSheet = ThisWorkbook.Sheets(SHEET_GAME)
    Set MenuSheet = ThisWorkbook.Sheets(SHEET_MENU)
    Set PauseSheet = ThisWorkbook.Sheets(SHEET_PAUSE)

    ' Collections
    Set Ducks = New Collection

    ' Game state
    GameRunning = False
    GamePaused = False
    GameEnded = False
    CurrentRound = 1

    ' Score
    Score = 0
    DucksShot = 0
    DucksMissed = 0

    ' Player
    Bullets = MaxBullets
    PlayerShot = False
    ReloadTime = 1#
    LastShotTime = Timer

    ' Ducks
    DucksPerRound = 5
    DucksSpawned = 0
    SpawnDelay = 1#
    LastSpawnTime = Timer

    ' Timing
    GameSpeed = 1#
    LastFrameTime = Timer

    ' Mouse
    MouseX = 0
    MouseY = 0

End Sub