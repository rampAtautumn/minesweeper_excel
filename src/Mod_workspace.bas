Attribute VB_Name = "mod_workspace"

Option Explicit

'====================================================
' WORKSPACE CONSTANTS
'====================================================

Private Const VIEW_PADDING_X As Long = 4
Private Const VIEW_PADDING_Y As Long = 6

Private Const MIN_ZOOM As Long = 55
Private Const MAX_ZOOM As Long = 130

'====================================================
' MAIN WORKSPACE ENTRY
'====================================================

Public Sub SetupWorkspace()

```
CleanupWorksheetView

ConfigureWorksheetSurface

CenterGameLayout

LockGameViewport
```

End Sub

'====================================================
' WORKSHEET CLEANUP
'====================================================

Public Sub CleanupWorksheetView()

```
With ActiveWindow

    .DisplayGridlines = False
    .DisplayHeadings = False
    .DisplayWorkbookTabs = False
    .Zoom = CalculateOptimalZoom()

End With

Application.DisplayFormulaBar = False
Application.DisplayStatusBar = False
```

End Sub

'====================================================
' WORKSHEET SURFACE
'====================================================

Private Sub ConfigureWorksheetSurface()

```
With GameSheet

    .Cells.RowHeight = TileSize
    .Cells.ColumnWidth = 2.8

    .Cells.Interior.Color = RGB(185, 185, 185)

End With
```

End Sub

'====================================================
' CENTER GAME
'====================================================

Public Sub CenterGameLayout()

```
Dim OffsetRow As Long
Dim OffsetCol As Long

CalculateViewportOffsets _
    OffsetRow, _
    OffsetCol

ApplyCenteredLayout _
    OffsetRow, _
    OffsetCol
```

End Sub

'====================================================
' OFFSET CALCULATION
'====================================================

Public Sub CalculateViewportOffsets( _
ByRef OffsetRow As Long, _
ByRef OffsetCol As Long _
)

```
Dim BoardWidth As Double
Dim BoardHeight As Double

BoardWidth = GetBoardPixelWidth()
BoardHeight = GetBoardPixelHeight()

OffsetCol = _
    Application.Max( _
        2, _
        Int((40 - BoardCols) / 2) _
    )

OffsetRow = _
    Application.Max( _
        3, _
        Int((25 - BoardRows) / 2) _
    )
```

End Sub

'====================================================
' APPLY CENTERED LAYOUT
'====================================================

Private Sub ApplyCenteredLayout( _
ByVal OffsetRow As Long, _
ByVal OffsetCol As Long _
)

```
BoardOriginRow = OffsetRow
BoardOriginCol = OffsetCol
```

End Sub

'====================================================
' LOCK VIEWPORT
'====================================================

Public Sub LockGameViewport()

```
Dim ScrollEndRow As Long
Dim ScrollEndCol As Long

ScrollEndRow = _
    BoardOriginRow + BoardRows + 8

ScrollEndCol = _
    BoardOriginCol + BoardCols + 8

GameSheet.ScrollArea = _
    GameSheet.Range( _
        GameSheet.Cells(1, 1), _
        GameSheet.Cells( _
            ScrollEndRow, _
            ScrollEndCol _
        ) _
    ).Address
```

End Sub

'====================================================
' UPDATE LAYOUT
'====================================================

Public Sub UpdateWorkspaceLayout()

```
CenterGameLayout

RealignBoardShapes

RealignHUD
```

End Sub

'====================================================
' BOARD BOUNDS
'====================================================

Public Function CalculateBoardBounds() As Range

```
Set CalculateBoardBounds = _
    GameSheet.Range( _
        GameSheet.Cells( _
            BoardOriginRow, _
            BoardOriginCol _
        ), _
        GameSheet.Cells( _
            BoardOriginRow + BoardRows - 1, _
            BoardOriginCol + BoardCols - 1 _
        ) _
    )
```

End Function

'====================================================
' HUD BOUNDS
'====================================================

Public Function CalculateHudBounds() As Range

```
Set CalculateHudBounds = _
    GameSheet.Range( _
        GameSheet.Cells( _
            BoardOriginRow - 2, _
            BoardOriginCol _
        ), _
        GameSheet.Cells( _
            BoardOriginRow - 1, _
            BoardOriginCol + BoardCols - 1 _
        ) _
    )
```

End Function

'====================================================
' PIXEL WIDTH
'====================================================

Public Function GetBoardPixelWidth() As Double

```
GetBoardPixelWidth = _
    BoardCols * TileSize
```

End Function

'====================================================
' PIXEL HEIGHT
'====================================================

Public Function GetBoardPixelHeight() As Double

```
GetBoardPixelHeight = _
    BoardRows * TileSize
```

End Function

'====================================================
' AUTO ZOOM
'====================================================

Private Function CalculateOptimalZoom() As Long

```
Dim ZoomValue As Long

If BoardCols >= 30 Then

    ZoomValue = 65

ElseIf BoardCols >= 16 Then

    ZoomValue = 80

Else

    ZoomValue = 110

End If

If ZoomValue < MIN_ZOOM Then
    ZoomValue = MIN_ZOOM
End If

If ZoomValue > MAX_ZOOM Then
    ZoomValue = MAX_ZOOM
End If

CalculateOptimalZoom = ZoomValue
```

End Function

'====================================================
' FRAME CREATION
'====================================================

Public Sub CreateBoardFrame()

```
Dim BoardArea As Range

SafeDeleteShape "game_frame"

Set BoardArea = CalculateBoardBounds()

With GameSheet.Shapes.AddShape( _
    msoShapeRectangle, _
    BoardArea.Left - 6, _
    BoardArea.Top - 6, _
    BoardArea.Width + 12, _
    BoardArea.Height + 12 _
)

    .Name = "game_frame"

    .Fill.ForeColor.RGB = RGB(120, 120, 120)

    .Line.ForeColor.RGB = RGB(70, 70, 70)

    .ZOrder msoSendToBack

End With
```

End Sub
Attribute VB_Name = "mod_workspace"

Option Explicit

'====================================================
' WORKSPACE CONSTANTS
'====================================================

Private Const VIEW_PADDING_X As Long = 4
Private Const VIEW_PADDING_Y As Long = 6

Private Const MIN_ZOOM As Long = 55
Private Const MAX_ZOOM As Long = 130

'====================================================
' MAIN WORKSPACE ENTRY
'====================================================

Public Sub SetupWorkspace()

```
CleanupWorksheetView

ConfigureWorksheetSurface

CenterGameLayout

LockGameViewport
```

End Sub

'====================================================
' WORKSHEET CLEANUP
'====================================================

Public Sub CleanupWorksheetView()

```
With ActiveWindow

    .DisplayGridlines = False
    .DisplayHeadings = False
    .DisplayWorkbookTabs = False
    .Zoom = CalculateOptimalZoom()

End With

Application.DisplayFormulaBar = False
Application.DisplayStatusBar = False
```

End Sub

'====================================================
' WORKSHEET SURFACE
'====================================================

Private Sub ConfigureWorksheetSurface()

```
With GameSheet

    .Cells.RowHeight = TileSize
    .Cells.ColumnWidth = 2.8

    .Cells.Interior.Color = RGB(185, 185, 185)

End With
```

End Sub

'====================================================
' CENTER GAME
'====================================================

Public Sub CenterGameLayout()

```
Dim OffsetRow As Long
Dim OffsetCol As Long

CalculateViewportOffsets _
    OffsetRow, _
    OffsetCol

ApplyCenteredLayout _
    OffsetRow, _
    OffsetCol
```

End Sub

'====================================================
' OFFSET CALCULATION
'====================================================

Public Sub CalculateViewportOffsets( _
ByRef OffsetRow As Long, _
ByRef OffsetCol As Long _
)

```
Dim BoardWidth As Double
Dim BoardHeight As Double

BoardWidth = GetBoardPixelWidth()
BoardHeight = GetBoardPixelHeight()

OffsetCol = _
    Application.Max( _
        2, _
        Int((40 - BoardCols) / 2) _
    )

OffsetRow = _
    Application.Max( _
        3, _
        Int((25 - BoardRows) / 2) _
    )
```

End Sub

'====================================================
' APPLY CENTERED LAYOUT
'====================================================

Private Sub ApplyCenteredLayout( _
ByVal OffsetRow As Long, _
ByVal OffsetCol As Long _
)

```
BoardOriginRow = OffsetRow
BoardOriginCol = OffsetCol
```

End Sub

'====================================================
' LOCK VIEWPORT
'====================================================

Public Sub LockGameViewport()

```
Dim ScrollEndRow As Long
Dim ScrollEndCol As Long

ScrollEndRow = _
    BoardOriginRow + BoardRows + 8

ScrollEndCol = _
    BoardOriginCol + BoardCols + 8

GameSheet.ScrollArea = _
    GameSheet.Range( _
        GameSheet.Cells(1, 1), _
        GameSheet.Cells( _
            ScrollEndRow, _
            ScrollEndCol _
        ) _
    ).Address
```

End Sub

'====================================================
' UPDATE LAYOUT
'====================================================

Public Sub UpdateWorkspaceLayout()

```
CenterGameLayout

RealignBoardShapes

RealignHUD
```

End Sub

'====================================================
' BOARD BOUNDS
'====================================================

Public Function CalculateBoardBounds() As Range

```
Set CalculateBoardBounds = _
    GameSheet.Range( _
        GameSheet.Cells( _
            BoardOriginRow, _
            BoardOriginCol _
        ), _
        GameSheet.Cells( _
            BoardOriginRow + BoardRows - 1, _
            BoardOriginCol + BoardCols - 1 _
        ) _
    )
```

End Function

'====================================================
' HUD BOUNDS
'====================================================

Public Function CalculateHudBounds() As Range

```
Set CalculateHudBounds = _
    GameSheet.Range( _
        GameSheet.Cells( _
            BoardOriginRow - 2, _
            BoardOriginCol _
        ), _
        GameSheet.Cells( _
            BoardOriginRow - 1, _
            BoardOriginCol + BoardCols - 1 _
        ) _
    )
```

End Function

'====================================================
' PIXEL WIDTH
'====================================================

Public Function GetBoardPixelWidth() As Double

```
GetBoardPixelWidth = _
    BoardCols * TileSize
```

End Function

'====================================================
' PIXEL HEIGHT
'====================================================

Public Function GetBoardPixelHeight() As Double

```
GetBoardPixelHeight = _
    BoardRows * TileSize
```

End Function

'====================================================
' AUTO ZOOM
'====================================================

Private Function CalculateOptimalZoom() As Long

```
Dim ZoomValue As Long

If BoardCols >= 30 Then

    ZoomValue = 65

ElseIf BoardCols >= 16 Then

    ZoomValue = 80

Else

    ZoomValue = 110

End If

If ZoomValue < MIN_ZOOM Then
    ZoomValue = MIN_ZOOM
End If

If ZoomValue > MAX_ZOOM Then
    ZoomValue = MAX_ZOOM
End If

CalculateOptimalZoom = ZoomValue
```

End Function

'====================================================
' FRAME CREATION
'====================================================

Public Sub CreateBoardFrame()

```
Dim BoardArea As Range

SafeDeleteShape "game_frame"

Set BoardArea = CalculateBoardBounds()

With GameSheet.Shapes.AddShape( _
    msoShapeRectangle, _
    BoardArea.Left - 6, _
    BoardArea.Top - 6, _
    BoardArea.Width + 12, _
    BoardArea.Height + 12 _
)

    .Name = "game_frame"

    .Fill.ForeColor.RGB = RGB(120, 120, 120)

    .Line.ForeColor.RGB = RGB(70, 70, 70)

    .ZOrder msoSendToBack

End With
```

End Sub
