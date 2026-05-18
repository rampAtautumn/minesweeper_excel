Attribute VB_Name = "mod_events"

Option Explicit

'====================================================
' INPUT ENTRY POINT
'====================================================

Public Sub HandleTileClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    RevealTile RowIndex, ColIndex

    RefreshBoard

End Sub

'====================================================
' RIGHT CLICK ENTRY
'====================================================

Public Sub HandleTileRightClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    ToggleFlag RowIndex, ColIndex

    RefreshBoard

End Sub

'====================================================
' DOUBLE CLICK ENTRY
'====================================================

Public Sub HandleTileDoubleClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    DoubleReveal RowIndex, ColIndex

    RefreshBoard

End Sub

'====================================================
' SHAPE CLICK ROUTER
'====================================================

Public Sub HandleShapeClick(ByVal ShapeName As String)

    Dim RowIndex As Long
    Dim ColIndex As Long

    If Not ParseTileShapeName( _
        ShapeName, _
        RowIndex, _
        ColIndex _
    ) Then

        Exit Sub

    End If

    HandleTileClick _
        RowIndex, _
        ColIndex

End Sub

'====================================================
' SHAPE RIGHT CLICK ROUTER
'====================================================

Public Sub HandleShapeRightClick(ByVal ShapeName As String)

    Dim RowIndex As Long
    Dim ColIndex As Long

    If Not ParseTileShapeName( _
        ShapeName, _
        RowIndex, _
        ColIndex _
    ) Then

        Exit Sub

    End If

    HandleTileRightClick _
        RowIndex, _
        ColIndex

End Sub

'====================================================
' SHAPE DOUBLE CLICK ROUTER
'====================================================

Public Sub HandleShapeDoubleClick(ByVal ShapeName As String)

    Dim RowIndex As Long
    Dim ColIndex As Long

    If Not ParseTileShapeName( _
        ShapeName, _
        RowIndex, _
        ColIndex _
    ) Then

        Exit Sub

    End If

    HandleTileDoubleClick _
        RowIndex, _
        ColIndex

End Sub

'====================================================
' TILE NAME PARSER
'====================================================

Public Function ParseTileShapeName( _
    ByVal ShapeName As String, _
    ByRef RowIndex As Long, _
    ByRef ColIndex As Long _
) As Boolean

    Dim Parts() As String

    ParseTileShapeName = False

    If Left$(ShapeName, Len(TILE_PREFIX)) <> TILE_PREFIX Then
        Exit Function
    End If

    Parts = Split(ShapeName, "_")

    If UBound(Parts) <> 2 Then
        Exit Function
    End If

    If Not IsNumeric(Parts(1)) Then
        Exit Function
    End If

    If Not IsNumeric(Parts(2)) Then
        Exit Function
    End If

    RowIndex = CLng(Parts(1))
    ColIndex = CLng(Parts(2))

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Function
    End If

    ParseTileShapeName = True

End Function

'====================================================
' MOUSE COORDINATE TRANSLATION
'====================================================

Public Function ScreenPositionToTile( _
    ByVal MouseX As Double, _
    ByVal MouseY As Double, _
    ByRef RowIndex As Long, _
    ByRef ColIndex As Long _
) As Boolean

    Dim BoardLeft As Double
    Dim BoardTop As Double

    Dim RelativeX As Double
    Dim RelativeY As Double

    ScreenPositionToTile = False

    BoardLeft = _
        GetTileCell(1, 1).Left

    BoardTop = _
        GetTileCell(1, 1).Top

    RelativeX = MouseX - BoardLeft
    RelativeY = MouseY - BoardTop

    If RelativeX < 0 Then Exit Function
    If RelativeY < 0 Then Exit Function

    ColIndex = Int(RelativeX / TileSize) + 1
    RowIndex = Int(RelativeY / TileSize) + 1

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Function
    End If

    ScreenPositionToTile = True

End Function

'====================================================
' HUD BUTTON ROUTING
'====================================================

Public Sub HandleRestartButton()

    StartNewGame

End Sub

'====================================================
' INPUT LOCKING
'====================================================

Public Function InputAllowed() As Boolean

    InputAllowed = _
        Not GameOver And _
        Not GameWon

End Function

'====================================================
' SAFE EVENT WRAPPERS
'====================================================

Public Sub SafeHandleLeftClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    On Error GoTo ErrorHandler

    HandleTileClick _
        RowIndex, _
        ColIndex

    Exit Sub

ErrorHandler:

    Debug.Print _
        "Left click error: " & _
        Err.Description

End Sub

Public Sub SafeHandleRightClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    On Error GoTo ErrorHandler

    HandleTileRightClick _
        RowIndex, _
        ColIndex

    Exit Sub

ErrorHandler:

    Debug.Print _
        "Right click error: " & _
        Err.Description

End Sub

Public Sub SafeHandleDoubleClick( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    On Error GoTo ErrorHandler

    HandleTileDoubleClick _
        RowIndex, _
        ColIndex

    Exit Sub

ErrorHandler:

    Debug.Print _
        "Double click error: " & _
        Err.Description

End Sub

'====================================================
' DEBUG UTILITIES
'====================================================

Public Sub DebugPrintTileFromShape(ByVal ShapeName As String)

    Dim r As Long
    Dim c As Long

    If ParseTileShapeName(ShapeName, r, c) Then

        Debug.Print _
            ShapeName & _
            " => Row: " & r & _
            " Col: " & c

    Else

        Debug.Print _
            "Invalid tile shape: " & ShapeName

    End If

End Sub