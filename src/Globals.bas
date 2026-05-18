Attribute VB_Name = "Globals"
Option Explicit

# mod_globals.bas

```vb
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
' CONSTANTS
'====================================================

Public Const TILE_PREFIX As String = "tile_"
Public Const HUD_PREFIX As String = "hud_"

Public Const SPRITE_HIDDEN As String = "hidden"
Public Const SPRITE_FLAG As String = "flag"
Public Const SPRITE_MINE As String = "mine"
Public Const SPRITE_ACTIVE_MINE As String = "active_mine"
Public Const SPRITE_EMPTY As String = "0"

