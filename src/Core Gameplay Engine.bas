Attribute VB_Name = " Engine"
Option Explicit

'====================================================
' INITIALIZATION
'====================================================

Public Sub InitializeGlobals()

    Set GameSheet = ThisWorkbook.Worksheets("Game")

    TileSize = 24

    BoardOriginRow = 5
    BoardOriginCol = 2

    RemainingFlags = 0

    GameStarted = False
    GameOver = False
    GameWon = False

    ExplodedRow = -1
    ExplodedCol = -1

    CurrentElapsedSeconds = 0

    TimerScheduled = False

End Sub

'====================================================
' BOARD MEMORY ALLOCATION
'====================================================

Public Sub AllocateBoardMemory()

    ReDim tablero(1 To BoardRows, 1 To BoardCols)
    ReDim revelado(1 To BoardRows, 1 To BoardCols)
    ReDim bandera(1 To BoardRows, 1 To BoardCols)

    ReDim DirtyTiles(1 To BoardRows, 1 To BoardCols)

    ReDim TileShapes(1 To BoardRows, 1 To BoardCols)

    ReDim LastRenderedSprite(1 To BoardRows, 1 To BoardCols)

End Sub

'====================================================
' RESET ARRAYS
'====================================================

Public Sub ResetBoardArrays()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            tablero(r, c) = 0
            revelado(r, c) = False
            bandera(r, c) = False
            DirtyTiles(r, c) = True

            LastRenderedSprite(r, c) = vbNullString

            Set TileShapes(r, c) = Nothing

        Next c

    Next r

End Sub

'====================================================
' TILE NAME HELPERS
'====================================================

Public Function GetTileShapeName(ByVal RowIndex As Long, ByVal ColIndex As Long) As String

    GetTileShapeName = TILE_PREFIX & RowIndex & "_" & ColIndex

End Function

'====================================================
' CELL HELPERS
'====================================================

Public Function GetTileCell(ByVal RowIndex As Long, ByVal ColIndex As Long) As Range

    Set GetTileCell = GameSheet.Cells(
        BoardOriginRow + RowIndex - 1,
        BoardOriginCol + ColIndex - 1
    )

End Function

'====================================================
' BOUNDS CHECKING
'====================================================

Public Function IsWithinBounds(ByVal RowIndex As Long, ByVal ColIndex As Long) As Boolean

    IsWithinBounds = _
        RowIndex >= 1 And _
        RowIndex <= BoardRows And _
        ColIndex >= 1 And _
        ColIndex <= BoardCols

End Function

'====================================================
' DIRTY TILE SYSTEM
'====================================================

Public Sub MarkTileDirty(ByVal RowIndex As Long, ByVal ColIndex As Long)

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    DirtyTiles(RowIndex, ColIndex) = True

End Sub

Public Sub MarkEntireBoardDirty()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            DirtyTiles(r, c) = True

        Next c

    Next r

End Sub

'====================================================
' DIFFICULTY CONFIGURATION
'====================================================

Public Sub ConfigureEasyMode()

    BoardRows = 8
    BoardCols = 8
    MineCount = 10

End Sub

Public Sub ConfigureMediumMode()

    BoardRows = 16
    BoardCols = 16
    MineCount = 40

End Sub

Public Sub ConfigureHardMode()

    BoardRows = 16
    BoardCols = 30
    MineCount = 99

End Sub
```

---

# mod_engine.bas

```vb
Attribute VB_Name = "mod_engine"

Option Explicit

'====================================================
' MAIN ENGINE ENTRY
'====================================================

Public Sub BootGame()

    Application.ScreenUpdating = False

    InitializeGlobals

    ConfigureEasyMode

    AllocateBoardMemory

    ResetBoardArrays

    InitializeEnvironment

    LoadAssets

    If Not VerifyAssets() Then

        MsgBox "Missing sprite assets.", vbCritical

        Application.ScreenUpdating = True

        Exit Sub

    End If

    InitializeBoard

    CreateBoardVisuals

    InitializeHUD

    RenderBoard

    StartGameTimer

    Application.ScreenUpdating = True

End Sub

'====================================================
' ENVIRONMENT SETUP
'====================================================

Public Sub InitializeEnvironment()

    With GameSheet

        .Cells.Clear

        .Activate

        ActiveWindow.DisplayGridlines = False

        ActiveWindow.DisplayHeadings = False

        ActiveWindow.Zoom = 100

    End With

    ConfigureBoardLayout

End Sub

'====================================================
' BOARD LAYOUT
'====================================================

Public Sub ConfigureBoardLayout()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        GameSheet.Rows(BoardOriginRow + r - 1).RowHeight = TileSize

    Next r

    For c = 1 To BoardCols

        GameSheet.Columns(BoardOriginCol + c - 1).ColumnWidth = TileSize / 5.3

    Next c

End Sub

'====================================================
' INITIAL BOARD CREATION
'====================================================

Public Sub InitializeBoard()

    Randomize

    GenerateMines

    CalculateAdjacentCounts

End Sub

'====================================================
' GAME START
'====================================================

Public Sub StartNewGame()

    ResetGame

    BootGame

End Sub

'====================================================
' GAME RESET
'====================================================

Public Sub ResetGame()

    Application.ScreenUpdating = False

    StopGameTimer

    ClearBoardSprites

    ClearHUD

    ResetBoardArrays

    GameOver = False
    GameWon = False
    GameStarted = False

    ExplodedRow = -1
    ExplodedCol = -1

    CurrentElapsedSeconds = 0

    Application.ScreenUpdating = True

End Sub

'====================================================
' GAME OVER FLOW
'====================================================

Public Sub HandleGameOver(ByVal MineRow As Long, ByVal MineCol As Long)

    If GameOver Then
        Exit Sub
    End If

    GameOver = True

    ExplodedRow = MineRow
    ExplodedCol = MineCol

    RevealAllMines

    MarkEntireBoardDirty

    RefreshBoard

    StopGameTimer

    MsgBox "Game Over", vbExclamation

End Sub

'====================================================
' GAME WIN FLOW
'====================================================

Public Sub HandleVictory()

    If GameWon Then
        Exit Sub
    End If

    GameWon = True

    GameOver = True

    StopGameTimer

    MsgBox "Victory", vbInformation

End Sub

'====================================================
' TIMER SYSTEM
'====================================================

Public Sub StartGameTimer()

    If TimerScheduled Then
        Exit Sub
    End If

    GameStartTime = Now

    ScheduleNextTimerTick

End Sub

Public Sub ScheduleNextTimerTick()

    If GameOver Then
        Exit Sub
    End If

    NextTimerTick = Now + TimeSerial(0, 0, 1)

    TimerScheduled = True

    Application.OnTime _
        EarliestTime:=NextTimerTick, _
        Procedure:="TimerTick", _
        Schedule:=True

End Sub

Public Sub TimerTick()

    TimerScheduled = False

    If GameOver Then
        Exit Sub
    End If

    CurrentElapsedSeconds = DateDiff("s", GameStartTime, Now)

    UpdateTimerHUD

    ScheduleNextTimerTick

End Sub

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

    StopGameTimer

    ClearBoardSprites

    ClearHUD

End Sub
```
