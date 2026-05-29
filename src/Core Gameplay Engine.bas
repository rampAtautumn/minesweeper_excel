Attribute VB_Name = "mod_engine"

Option Explicit

'====================================================
' MAIN ENGINE ENTRY
'====================================================

Public Sub BootGame()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    '================================================
    ' CORE INITIALIZATION
    '================================================

    InitializeGlobals

    '================================================
    ' DEFAULT DIFFICULTY
    '================================================

    If CurrentDifficulty = 0 Then

        SetEasyDifficulty

    End If

    '================================================
    ' MEMORY
    '================================================

    ReallocateGameMemory

    ResetBoardArrays

    '================================================
    ' ASSET SYSTEM
    '================================================

    LoadAssets

    If Not VerifyAssets() Then

        MsgBox _
            "Missing sprite assets.", _
            vbCritical

        GoTo Cleanup

    End If

    '================================================
    ' ENVIRONMENT
    '================================================

    InitializeEnvironment

    SetupWorkspace

    ConfigureBoardLayout

    ApplyClassicWindowStyle

    '================================================
    ' GAME STATE
    '================================================

    InitializeBoard

    RemainingFlags = MineCount

    CurrentElapsedSeconds = 0

    GameStarted = False
    GameOver = False
    GameWon = False

    ExplodedRow = -1
    ExplodedCol = -1

    '================================================
    ' VISUAL INITIALIZATION
    '================================================

    CreateBoardVisuals

    InitializeHUD

    CreateDifficultyButtons

    MarkEntireBoardDirty

    RenderBoard

    RefreshHUD

    '================================================
    ' FORCE COMPLETE REDRAW
    '================================================

    MarkEntireBoardDirty

    RenderBoard

    RefreshHUD
    
    ScrollToGame

Cleanup:

    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "BootGame Error:" & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub

'====================================================
' ENVIRONMENT SETUP
'====================================================

Public Sub InitializeEnvironment()

    With GameSheet

        .Cells.RowHeight = TileSize

        .Cells.ColumnWidth = 2.8

        .Cells.Interior.Color = _
            RGB(192, 192, 192)

    End With

End Sub

'====================================================
' BOARD LAYOUT
'====================================================

Public Sub ConfigureBoardLayout()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        GameSheet.Rows( _
            BoardOriginRow + r - 1 _
        ).RowHeight = TileSize

    Next r

    For c = 1 To BoardCols

        GameSheet.Columns( _
            BoardOriginCol + c - 1 _
        ).ColumnWidth = TileSize / 5.3

    Next c

End Sub

'====================================================
' BOARD INITIALIZATION
'====================================================

Public Sub InitializeBoard()

    Randomize

    InitializeEmptyBoard

    GenerateMines

    CalculateAdjacentCounts

End Sub

'====================================================
' START GAME SESSION
'====================================================

Public Sub StartGameSession()

    If GameStarted Then
        Exit Sub
    End If

    GameStarted = True

    StartGameTimer

End Sub

'====================================================
' START NEW GAME
'====================================================
Public Sub StartNewGame()

    ResetGame

    BootGame

End Sub

'====================================================
' FULL GAME RESET
'====================================================

Public Sub ResetGame()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    '================================================
    ' STOP ACTIVE SYSTEMS
    '================================================

    StopGameTimer

    ClearHoverEffect

    '================================================
    ' CLEAR VISUALS
    '================================================

    ClearBoardSprites

    ClearHUD

    ClearDifficultyButtons

    '================================================
    ' RESET GAME STATE
    '================================================

    GameStarted = False

    GameOver = False

    GameWon = False

    RemainingFlags = MineCount

    CurrentElapsedSeconds = 0

    ExplodedRow = -1
    ExplodedCol = -1

    '================================================
    ' RESET MEMORY
    '================================================
    
    ReallocateGameMemory
    
    ResetBoardArrays
    
    '================================================
    ' REBUILD BOARD
    '================================================
    
    InitializeBoard
    
    '================================================
    ' RECALCULATE LAYOUT
    '================================================
    
    SetupWorkspace
    
    ConfigureBoardLayout

    '================================================
    ' REBUILD VISUALS
    '================================================

    CreateBoardFrame

    CreateBoardVisuals

    InitializeHUD

    CreateDifficultyButtons

    '================================================
    ' FORCE FULL REDRAW
    '================================================

    MarkEntireBoardDirty

    RenderBoard

    RefreshHUD

    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "ResetGame Error:" & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub

'====================================================
' HANDLE GAME OVER
'====================================================

Public Sub HandleGameOver( _
    ByVal MineRow As Long, _
    ByVal MineCol As Long _
)

    If GameOver Then
        Exit Sub
    End If

    GameOver = True

    ExplodedRow = MineRow
    ExplodedCol = MineCol

    RevealAllMines

    MarkEntireBoardDirty

    PlayLossEffect

    RefreshBoard

    StopGameTimer

End Sub

'====================================================
' HANDLE VICTORY
'====================================================

Public Sub HandleVictory()

    If GameWon Then
        Exit Sub
    End If

    GameWon = True
    GameOver = True

    StopGameTimer

    PlayVictoryEffect

    RefreshHUD

End Sub

'====================================================
' REVEAL ALL MINES
'====================================================

Public Sub RevealAllMines()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) = -1 Then

                revelado(r, c) = True

                DirtyTiles(r, c) = True

            End If

        Next c

    Next r

End Sub

'====================================================
' TIMER START
'====================================================

Public Sub StartGameTimer()

    If TimerScheduled Then
        Exit Sub
    End If

    GameStartTime = Now

    ScheduleNextTimerTick

End Sub

'====================================================
' TIMER SCHEDULING
'====================================================

Public Sub ScheduleNextTimerTick()

    If GameOver Then
        Exit Sub
    End If

    NextTimerTick = _
        Now + TimeSerial(0, 0, 1)

    TimerScheduled = True

    Application.OnTime _
        EarliestTime:=NextTimerTick, _
        Procedure:="TimerTick", _
        Schedule:=True

End Sub

'====================================================
' TIMER TICK
'====================================================

Public Sub TimerTick()

    TimerScheduled = False

    If GameOver Then
        Exit Sub
    End If

    If Not GameStarted Then
        Exit Sub
    End If

    CurrentElapsedSeconds = _
        DateDiff("s", GameStartTime, Now)
    
    RefreshHUD
    
    ScheduleNextTimerTick

End Sub

'====================================================
' TIMER STOP
'====================================================

Public Sub StopGameTimer()

    On Error Resume Next

    If TimerScheduled Then

        Application.OnTime _
            EarliestTime:=NextTimerTick, _
            Procedure:="TimerTick", _
            Schedule:=False

    End If

    TimerScheduled = False

    On Error GoTo 0

End Sub

'====================================================
' SAFE SHUTDOWN
'====================================================

Public Sub ShutdownGame()

    On Error Resume Next

    StopGameTimer

    ClearHoverEffect

    ClearBoardSprites

    ClearHUD

    ClearDifficultyButtons

    GameSheet.ScrollArea = vbNullString

    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True

    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True

    GameStarted = False
    GameOver = False
    GameWon = False

    On Error GoTo 0

End Sub

'====================================================
' FULL REFRESH
'====================================================

Public Sub RefreshEntireGame()

    RefreshBoard

    RefreshHUD

End Sub

'====================================================
' HARD REFRESH
'====================================================

Public Sub HardRefresh()

    Application.ScreenUpdating = False

    MarkEntireBoardDirty

    RefreshBoard

    RefreshHUD

    Application.ScreenUpdating = True

End Sub
Public Sub RebuildVisualLayer()

    ClearBoardSprites

    ClearHUD

    CreateBoardVisuals

    InitializeHUD

    MarkEntireBoardDirty

    RenderBoard

    RefreshHUD

End Sub
Public Sub ScrollToGame()

    Application.Goto _
        GameSheet.Cells( _
            BoardOriginRow, _
            BoardOriginCol _
        ), _
        True

End Sub

'====================================================
' DEBUG ENGINE STATE
'====================================================

Public Sub DebugPrintEngineState()

    Debug.Print _
        "===== ENGINE STATE ====="

    Debug.Print _
        "GameStarted: " & GameStarted

    Debug.Print _
        "GameOver: " & GameOver

    Debug.Print _
        "GameWon: " & GameWon

    Debug.Print _
        "BoardRows: " & BoardRows

    Debug.Print _
        "BoardCols: " & BoardCols

    Debug.Print _
        "MineCount: " & MineCount

    Debug.Print _
        "RemainingFlags: " & RemainingFlags

    Debug.Print _
        "ElapsedSeconds: " & _
        CurrentElapsedSeconds

End Sub
