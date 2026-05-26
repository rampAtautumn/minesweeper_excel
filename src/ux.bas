Attribute VB_Name = "mod_ux"

Option Explicit

'====================================================
' HOVER STATE
'====================================================

Private HoverShapeName As String

'====================================================
' CLEAR HOVER
'====================================================

Public Sub ClearHoverEffect()

    On Error Resume Next

    If Len(HoverShapeName) > 0 Then

        GameSheet.Shapes(HoverShapeName).Line.Visible = msoFalse

    End If

    HoveredRow = -1
    HoveredCol = -1

    HoverShapeName = vbNullString

    On Error GoTo 0

End Sub

'====================================================
' APPLY HOVER EFFECT
'====================================================

Public Sub ApplyHoverEffect( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim TileShape As Shape

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    ClearHoverEffect

    If TileShapes(RowIndex, ColIndex) Is Nothing Then
        Exit Sub
    End If

    Set TileShape = TileShapes(RowIndex, ColIndex)

    With TileShape.Line

        .Visible = msoTrue

        .Weight = 1.5

        .ForeColor.RGB = RGB(255, 255, 255)

    End With

    HoveredRow = RowIndex
    HoveredCol = ColIndex

    HoverShapeName = TileShape.Name

End Sub

'====================================================
' PRESSED TILE EFFECT
'====================================================

Public Sub ApplyPressedTileEffect( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If TileShapes(RowIndex, ColIndex) Is Nothing Then
        Exit Sub
    End If

    With TileShapes(RowIndex, ColIndex)

        .IncrementLeft 1
        .IncrementTop 1

    End With

End Sub

'====================================================
' RELEASE TILE EFFECT
'====================================================

Public Sub ReleasePressedTileEffect( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If TileShapes(RowIndex, ColIndex) Is Nothing Then
        Exit Sub
    End If

    With TileShapes(RowIndex, ColIndex)

        .IncrementLeft -1
        .IncrementTop -1

    End With

End Sub

'====================================================
' LOSS EFFECT
'====================================================

Public Sub PlayLossEffect()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) = -1 Then

                DirtyTiles(r, c) = True

            End If

        Next c

    Next r

End Sub

'====================================================
' VICTORY EFFECT
'====================================================

Public Sub PlayVictoryEffect()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) = -1 Then

                bandera(r, c) = True

                DirtyTiles(r, c) = True

            End If

        Next c

    Next r

    RefreshBoard

End Sub

'====================================================
' RESTART BUTTON HOVER
'====================================================

Public Sub ApplyRestartHover()

    On Error Resume Next

    With GameSheet.Shapes("restart_button")

        .Line.Visible = msoTrue

        .Line.ForeColor.RGB = RGB(255, 255, 255)

        .Line.Weight = 2

    End With

    On Error GoTo 0

End Sub

'====================================================
' CLEAR RESTART HOVER
'====================================================

Public Sub ClearRestartHover()

    On Error Resume Next

    With GameSheet.Shapes("restart_button")

        .Line.Visible = msoFalse

    End With

    On Error GoTo 0

End Sub

'====================================================
' RESTART PRESSED EFFECT
'====================================================

Public Sub ApplyRestartPressedEffect()

    On Error Resume Next

    With GameSheet.Shapes("restart_button")

        .IncrementLeft 1

        .IncrementTop 1

    End With

    On Error GoTo 0

End Sub

'====================================================
' RESTART RELEASE EFFECT
'====================================================

Public Sub ReleaseRestartPressedEffect()

    On Error Resume Next

    With GameSheet.Shapes("restart_button")

        .IncrementLeft -1

        .IncrementTop -1

    End With

    On Error GoTo 0

End Sub