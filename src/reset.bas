Attribute VB_Name = "mod_reset"

Option Explicit

'====================================================
' FULL GAME RESET
'====================================================

Public Sub ResetGame()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    StopGameTimer

    ResetGameState

    ClearVisualState

    ResetBoardData

    ResetRenderSystems

    ResetHudState

    RebuildGameSession

    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    MsgBox _
        "Reset failed:" & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub

'====================================================
' GAME STATE RESET
'====================================================

Private Sub ResetGameState()

    GameStarted = False

    GameOver = False

    GameWon = False

    RemainingFlags = MineCount

    CurrentElapsedSeconds = 0

    ExplodedRow = -1
    ExplodedCol = -1

End Sub

'====================================================
' VISUAL CLEANUP
'====================================================

Private Sub ClearVisualState()

    ClearBoardSprites

    ClearHUD

    RemoveOrphanShapes

End Sub

'====================================================
' ARRAY RESET
'====================================================

Private Sub ResetBoardData()

    ResetBoardArrays

    InitializeEmptyBoard

    GenerateMines

    CalculateAdjacentCounts

End Sub

'====================================================
' CACHE RESET
'====================================================

Private Sub ResetRenderSystems()

    ResetRenderCache

    MarkEntireBoardDirty

End Sub

'====================================================
' HUD RESET
'====================================================

Private Sub ResetHudState()

    InitializeHUD
    RefreshHUD

End Sub

'====================================================
' SESSION REBUILD
'====================================================

Private Sub RebuildGameSession()

    CreateBoardVisuals

    RenderBoard

    StartGameTimer

End Sub

'====================================================
' HARD RESET
'====================================================

Public Sub HardResetGame()

    On Error Resume Next

    Application.ScreenUpdating = False

    StopGameTimer

    ClearBoardSprites

    ClearHUD

    Erase tablero
    Erase revelado
    Erase bandera

    Erase DirtyTiles

    Erase TileShapes

    Erase LastRenderedSprite

    Set SpritePaths = Nothing

    GameStarted = False
    GameOver = False
    GameWon = False

    HudInitialized = False

    ExplodedRow = -1
    ExplodedCol = -1

    CurrentElapsedSeconds = 0

    RemainingFlags = 0

    Application.ScreenUpdating = True

    On Error GoTo 0

End Sub

'====================================================
' ORPHAN SHAPE CLEANUP
'====================================================

Public Sub RemoveOrphanShapes()

    Dim shp As Shape

    Dim ShapesToDelete As Collection

    Dim Item As Variant

    Set ShapesToDelete = New Collection

    For Each shp In GameSheet.Shapes

        If IsBoardShape(shp.Name) Or _
           IsHudShape(shp.Name) Then

            ShapesToDelete.Add shp.Name

        End If

    Next shp

    For Each Item In ShapesToDelete

        SafeDeleteShape CStr(Item)

    Next Item

End Sub

'====================================================
' BOARD SHAPE DETECTION
'====================================================

Private Function IsBoardShape( _
    ByVal ShapeName As String _
) As Boolean

    IsBoardShape = _
        Left$(ShapeName, Len(TILE_PREFIX)) = _
        TILE_PREFIX

End Function

'====================================================
' HUD SHAPE DETECTION
'====================================================

Private Function IsHudShape( _
    ByVal ShapeName As String _
) As Boolean

    IsHudShape = _
        Left$(ShapeName, Len(HUD_PREFIX)) = _
        HUD_PREFIX

End Function

'====================================================
' SAFE ENGINE RESTART
'====================================================

Public Sub RestartEngine()

    Application.ScreenUpdating = False

    StopGameTimer

    ClearHoverEffect

    ClearBoardSprites

    ClearHUD

    '============================================
    ' DESTROY OLD MEMORY
    '============================================

    Erase tablero
    Erase revelado
    Erase bandera

    Erase DirtyTiles
    Erase TileShapes
    Erase LastRenderedSprite

    '============================================
    ' REALLOCATE NEW MEMORY
    '============================================

    AllocateBoardMemory

    '============================================
    ' RESET DATA
    '============================================

    InitializeBoard

    RemainingFlags = MineCount

    CurrentElapsedSeconds = 0

    GameStarted = False
    GameOver = False
    GameWon = False

    ExplodedRow = -1
    ExplodedCol = -1

    '============================================
    ' REBUILD VISUALS
    '============================================

    RebuildVisualLayer

    Application.ScreenUpdating = True

End Sub

'====================================================
' NEW DIFFICULTY RESTART
'====================================================

Public Sub RestartEasy()

    HardResetGame

    SetEasyDifficulty


    BootGame

End Sub

Public Sub RestartMedium()

    HardResetGame

    SetMediumDifficulty

    BootGame

End Sub

Public Sub RestartHard()

    HardResetGame

    SetHardDifficulty

    BootGame

End Sub

'====================================================
' TIMER RESET
'====================================================

Public Sub ResetTimerState()

    StopGameTimer

    CurrentElapsedSeconds = 0

    TimerScheduled = False

End Sub

'====================================================
' SAFE ARRAY REALLOCATION
'====================================================

Public Sub ReallocateGameMemory()

    Erase tablero
    Erase revelado
    Erase bandera

    Erase DirtyTiles

    Erase TileShapes

    Erase LastRenderedSprite

    AllocateBoardMemory

End Sub

'====================================================
' FULL ENVIRONMENT RESET
'====================================================

Public Sub ResetWorksheetEnvironment()

    With GameSheet

        .Cells.Clear

        .DrawingObjects.Delete

    End With

End Sub

'====================================================
' RESET VALIDATION
'====================================================

Public Function ValidateResetState() As Boolean

    ValidateResetState = False

    If GameOver Then Exit Function

    If GameWon Then Exit Function

    If CurrentElapsedSeconds <> 0 Then Exit Function

    If RemainingFlags <> MineCount Then Exit Function

    ValidateResetState = True

End Function

'====================================================
' DEBUG UTILITIES
'====================================================

Public Sub DebugPrintResetState()

    Debug.Print _
        "===== RESET STATE ====="

    Debug.Print _
        "GameStarted: " & GameStarted

    Debug.Print _
        "GameOver: " & GameOver

    Debug.Print _
        "GameWon: " & GameWon

    Debug.Print _
        "RemainingFlags: " & RemainingFlags

    Debug.Print _
        "CurrentElapsedSeconds: " & _
        CurrentElapsedSeconds

    Debug.Print _
        "HudInitialized: " & HudInitialized

End Sub
