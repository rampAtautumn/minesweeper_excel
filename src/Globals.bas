Attribute VB_Name = "Globals"
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
' SHEET MANAGEMENT
'=======================

Public Function SheetExists(ByVal sheetName As String) As Boolean
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = (Not ws Is Nothing)
    
End Function

Public Sub EnsureSheet(ByVal sheetName As String)
    If Not SheetExists(sheetName) Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
End Sub

'=======================
' CLEAN WORKBOOK
'=======================

Public Sub CleanWorkbook()

    Dim ws As Worksheet
    Dim keepNames As Object
    Dim sheetCount As Long
    
    Set keepNames = CreateObject("Scripting.Dictionary")
    
    keepNames.Add SHEET_GAME, True
    keepNames.Add SHEET_MENU, True
    keepNames.Add SHEET_PAUSE, True
    
    Application.DisplayAlerts = False
    
    ' Cache count before loop
    sheetCount = ThisWorkbook.Worksheets.Count
    
    ' Iterate backward to avoid collection invalidation during deletion
    Dim i As Long
    For i = sheetCount To 1 Step -1
        
        If ThisWorkbook.Worksheets.Count <= 1 Then Exit For
        
        Set ws = ThisWorkbook.Worksheets(i)
        
        If Not keepNames.exists(ws.Name) Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
        
    Next i
    
    Application.DisplayAlerts = True

End Sub

'=======================
' INITIALIZATION
'=======================

Public Sub InitializeGlobals()

    '-----------------------
    ' 1. CLEAN ENVIRONMENT
    '-----------------------
    CleanWorkbook

    '-----------------------
    ' 2. ENSURE SHEETS
    '-----------------------
    EnsureSheet SHEET_GAME
    EnsureSheet SHEET_MENU
    EnsureSheet SHEET_PAUSE


    '-----------------------
    ' 3. ASSIGN REFERENCES
    '-----------------------
    On Error Resume Next
    Set GameSheet = ThisWorkbook.Sheets(SHEET_GAME)
    Set MenuSheet = ThisWorkbook.Sheets(SHEET_MENU)
    Set PauseSheet = ThisWorkbook.Sheets(SHEET_PAUSE)
    On Error GoTo 0
    
    ' Validate sheet references were assigned
    If GameSheet Is Nothing Or MenuSheet Is Nothing Or PauseSheet Is Nothing Then
        Debug.Print "❌ ERROR: Failed to initialize sheet references"
        Exit Sub
    End If
    
    '-----------------------
    ' 4. COLLECTIONS
    '-----------------------
    Set Ducks = New Collection

    '-----------------------
    ' 5. GAME STATE
    '-----------------------
    GameRunning = False
    GamePaused = False
    GameEnded = False
    CurrentRound = 1

    '-----------------------
    ' 6. SCORE
    '-----------------------
    Score = 0
    DucksShot = 0
    DucksMissed = 0

    '-----------------------
    ' 7. PLAYER
    '-----------------------
    Bullets = MaxBullets
    PlayerShot = False
    ReloadTime = 1#
    LastShotTime = Timer

    '-----------------------
    ' 8. DUCKS
    '-----------------------
    DucksPerRound = 5
    DucksSpawned = 0
    SpawnDelay = 1#
    LastSpawnTime = Timer

    '-----------------------
    ' 9. TIMING
    '-----------------------
    GameSpeed = 1#
    LastFrameTime = Timer
    DeltaTime = 0#

    '-----------------------
    ' 10. MOUSE
    '-----------------------
    MouseX = 0
    MouseY = 0

End Sub