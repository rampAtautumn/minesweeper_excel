Attribute VB_Name = "mod_globals"

Option Explicit

'====================================================
' BOARD STATE ARRAYS
'====================================================

Public tablero() As Integer
Public revelado() As Boolean
Public bandera() As Boolean
Public DirtyTiles() As Boolean

'====================================================
' SHAPE CACHE
'====================================================

Public TileShapes() As Shape

'====================================================
' WORKSHEET REFERENCES
'====================================================

Public GameSheet As Worksheet

'====================================================
' BOARD CONFIGURATION
'====================================================

Public BoardRows As Long
Public BoardCols As Long
Public MineCount As Long

Public TileSize As Double

Public BoardOriginRow As Long
Public BoardOriginCol As Long


'====================================================
' GAME STATE
'====================================================

Public GameStarted As Boolean
Public GameOver As Boolean
Public GameWon As Boolean

Public RemainingFlags As Long

Public ExplodedRow As Long
Public ExplodedCol As Long

'====================================================
' TIMER SYSTEM
'====================================================

Public GameStartTime As Date
Public CurrentElapsedSeconds As Long

Public TimerScheduled As Boolean
Public NextTimerTick As Date

'====================================================
' ASSET REGISTRY
'====================================================

Public SpritePaths As Object

Public AssetsRoot As String

'====================================================
' RENDER CACHE
'====================================================

Public LastRenderedSprite() As String

'====================================================
' HUD REFERENCES
'====================================================

Public HudInitialized As Boolean

'====================================================
' UX STATE
'====================================================

Public HoveredRow As Long
Public HoveredCol As Long

'====================================================
' CONSTANTS
'====================================================

Public Const TILE_PREFIX As String = "tile_"
Public Const HUD_PREFIX As String = "hud_"

Public Const SPRITE_HIDDEN As String = "hidden"
Public Const SPRITE_FLAG As String = "flag"
Public Const SPRITE_MINE As String = "mine"
Public Const SPRITE_ACTIVE_MINE As String = "active_mine"
Public Const SPRITE_EMPTY As String = "0"

'====================================================
' INITIALIZATION
'====================================================

Public Sub InitializeGlobals()

    Set GameSheet = _
        ThisWorkbook.Worksheets("Game")

    TileSize = 24

    BoardOriginRow = 5
    BoardOriginCol = 2

    RemainingFlags = 0

    GameStarted = False
    GameOver = False
    GameWon = False

    ExplodedRow = -1
    ExplodedCol = -1

    HoveredRow = -1
    HoveredCol = -1

    CurrentElapsedSeconds = 0

    TimerScheduled = False

End Sub

'====================================================
' MEMORY ALLOCATION
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
' BOUNDS CHECKING
'====================================================

Public Function IsWithinBounds( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Boolean

    IsWithinBounds = _
        RowIndex >= 1 And _
        RowIndex <= BoardRows And _
        ColIndex >= 1 And _
        ColIndex <= BoardCols

End Function

'====================================================
' TILE CELL HELPER
'====================================================

Public Function GetTileCell( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Range

    Set GetTileCell = _
        GameSheet.Cells( _
            BoardOriginRow + RowIndex - 1, _
            BoardOriginCol + ColIndex - 1 _
        )

End Function

'====================================================
' TILE SHAPE NAME
'====================================================

Public Function GetTileShapeName( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As String

    GetTileShapeName = _
        TILE_PREFIX & _
        RowIndex & "_" & _
        ColIndex

End Function

'====================================================
' DIRTY TILE HELPERS
'====================================================

Public Sub MarkTileDirty( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

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

