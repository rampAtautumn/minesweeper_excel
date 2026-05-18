Attribute VB_Name = "mod_render"

Option Explicit

'====================================================
' FULL BOARD RENDER
'====================================================

Public Sub RenderBoard()

    Dim r As Long
    Dim c As Long

    Application.ScreenUpdating = False

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            RenderTile r, c

        Next c

    Next r

    Application.ScreenUpdating = True

End Sub

'====================================================
' DIRTY TILE REFRESH
'====================================================

Public Sub RefreshBoard()

    Dim r As Long
    Dim c As Long

    Application.ScreenUpdating = False

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If DirtyTiles(r, c) Then

                RenderTile r, c

                DirtyTiles(r, c) = False

            End If

        Next c

    Next r

    Application.ScreenUpdating = True

End Sub

'====================================================
' TILE RENDERING
'====================================================

Public Sub RenderTile( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim SpriteKey As String

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    SpriteKey = ResolveTileSprite( _
        RowIndex, _
        ColIndex _
    )

    If LastRenderedSprite(RowIndex, ColIndex) = SpriteKey Then

        Exit Sub

    End If

    UpdateTileSprite _
        RowIndex, _
        ColIndex, _
        SpriteKey

    LastRenderedSprite(RowIndex, ColIndex) = _
        SpriteKey

End Sub

'====================================================
' INITIAL VISUAL CREATION
'====================================================

Public Sub CreateBoardVisuals()

    Dim r As Long
    Dim c As Long

    Application.ScreenUpdating = False

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            CreateTileShape r, c

        Next c

    Next r

    Application.ScreenUpdating = True

End Sub

'====================================================
' TILE SHAPE CREATION
'====================================================

Private Sub CreateTileShape( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim TileCell As Range

    Dim ShapeName As String

    Dim NewShape As Shape

    ShapeName = _
        GetTileShapeName( _
            RowIndex, _
            ColIndex _
        )

    Set TileCell = _
        GetTileCell( _
            RowIndex, _
            ColIndex _
        )

    SafeDeleteShape ShapeName

    Set NewShape = _
        GameSheet.Shapes.AddPicture( _
            Filename:=GetSpritePath("hidden"), _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=TileCell.Left, _
            Top:=TileCell.Top, _
            Width:=TileSize, _
            Height:=TileSize _
        )

    With NewShape

        .Name = ShapeName

        .Placement = xlMoveAndSize

        .LockAspectRatio = msoFalse

        .OnAction = "TileShapeClicked"

    End With

    Set TileShapes(RowIndex, ColIndex) = _
        NewShape

    LastRenderedSprite(RowIndex, ColIndex) = _
        vbNullString

End Sub

'====================================================
' TILE SPRITE UPDATE
'====================================================

Public Sub UpdateTileSprite( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long, _
    ByVal SpriteKey As String _
)

    Dim TileCell As Range

    Dim ShapeName As String

    Dim ExistingShape As Shape
    Dim NewShape As Shape

    ShapeName = _
        GetTileShapeName( _
            RowIndex, _
            ColIndex _
        )

    Set TileCell = _
        GetTileCell( _
            RowIndex, _
            ColIndex _
        )

    SafeDeleteShape ShapeName

    Set NewShape = _
        GameSheet.Shapes.AddPicture( _
            Filename:=GetSpritePath(SpriteKey), _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=TileCell.Left, _
            Top:=TileCell.Top, _
            Width:=TileSize, _
            Height:=TileSize _
        )

    With NewShape

        .Name = ShapeName

        .Placement = xlMoveAndSize

        .LockAspectRatio = msoFalse

        .OnAction = "TileShapeClicked"

    End With

    Set TileShapes(RowIndex, ColIndex) = _
        NewShape

End Sub

'====================================================
' SHAPE CLEANUP
'====================================================

Public Sub ClearBoardSprites()

    Dim r As Long
    Dim c As Long

    Dim ShapeName As String

    Application.ScreenUpdating = False

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            ShapeName = _
                GetTileShapeName(r, c)

            SafeDeleteShape ShapeName

            Set TileShapes(r, c) = Nothing

        Next c

    Next r

    Application.ScreenUpdating = True

End Sub

'====================================================
' SAFE SHAPE DELETION
'====================================================

Public Sub SafeDeleteShape( _
    ByVal ShapeName As String _
)

    On Error Resume Next

    GameSheet.Shapes(ShapeName).Delete

    On Error GoTo 0

End Sub

'====================================================
' FORCE FULL REDRAW
'====================================================

Public Sub ForceFullRender()

    MarkEntireBoardDirty

    RefreshBoard

End Sub

'====================================================
' SHAPE ACTION ROUTER
'====================================================

Public Sub TileShapeClicked()

    Dim ShapeName As String

    ShapeName = _
        Application.Caller

    HandleShapeClick ShapeName

End Sub

'====================================================
' VISUAL ALIGNMENT
'====================================================

Public Sub RealignBoardShapes()

    Dim r As Long
    Dim c As Long

    Dim TileCell As Range

    If TileShapes Is Nothing Then
        Exit Sub
    End If

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If Not TileShapes(r, c) Is Nothing Then

                Set TileCell = _
                    GetTileCell(r, c)

                With TileShapes(r, c)

                    .Left = TileCell.Left
                    .Top = TileCell.Top
                    .Width = TileSize
                    .Height = TileSize

                End With

            End If

        Next c

    Next r

End Sub