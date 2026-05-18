Attribute VB_Name = "mod_render_cache"

Option Explicit

'====================================================
' CACHE RESET
'====================================================

Public Sub ResetRenderCache()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            LastRenderedSprite(r, c) = _
                vbNullString

            DirtyTiles(r, c) = True

        Next c

    Next r

End Sub

'====================================================
' TILE CACHE INVALIDATION
'====================================================

Public Sub InvalidateTileCache( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    LastRenderedSprite(RowIndex, ColIndex) = _
        vbNullString

    DirtyTiles(RowIndex, ColIndex) = True

End Sub

'====================================================
' MASS CACHE INVALIDATION
'====================================================

Public Sub InvalidateEntireCache()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            LastRenderedSprite(r, c) = _
                vbNullString

            DirtyTiles(r, c) = True

        Next c

    Next r

End Sub

'====================================================
' CACHE SYNCHRONIZATION
'====================================================

Public Sub SynchronizeRenderCache()

    Dim r As Long
    Dim c As Long

    Dim SpriteKey As String

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            SpriteKey = _
                ResolveTileSprite(r, c)

            LastRenderedSprite(r, c) = _
                SpriteKey

        Next c

    Next r

End Sub

'====================================================
' DIRTY TILE COUNT
'====================================================

Public Function CountDirtyTiles() As Long

    Dim r As Long
    Dim c As Long

    CountDirtyTiles = 0

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If DirtyTiles(r, c) Then

                CountDirtyTiles = _
                    CountDirtyTiles + 1

            End If

        Next c

    Next r

End Function

'====================================================
' CACHE DEBUG
'====================================================

Public Sub DebugPrintRenderCache()

    Dim r As Long
    Dim c As Long

    Dim OutputLine As String

    Debug.Print _
        "===== RENDER CACHE ====="

    For r = 1 To BoardRows

        OutputLine = vbNullString

        For c = 1 To BoardCols

            OutputLine = _
                OutputLine & _
                "[" & _
                LastRenderedSprite(r, c) & _
                "] "

        Next c

        Debug.Print OutputLine

    Next r

End Sub

'====================================================
' DIRTY TILE DEBUG
'====================================================

Public Sub DebugPrintDirtyTiles()

    Dim r As Long
    Dim c As Long

    Dim OutputLine As String

    Debug.Print _
        "===== DIRTY TILES ====="

    For r = 1 To BoardRows

        OutputLine = vbNullString

        For c = 1 To BoardCols

            If DirtyTiles(r, c) Then

                OutputLine = _
                    OutputLine & " 1"

            Else

                OutputLine = _
                    OutputLine & " 0"

            End If

        Next c

        Debug.Print OutputLine

    Next r

End Sub

'====================================================
' CACHE VALIDATION
'====================================================

Public Function TileCacheMatchesState( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Boolean

    Dim ExpectedSprite As String

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        TileCacheMatchesState = False

        Exit Function

    End If

    ExpectedSprite = _
        ResolveTileSprite( _
            RowIndex, _
            ColIndex _
        )

    TileCacheMatchesState = _
        (LastRenderedSprite(RowIndex, ColIndex) = _
            ExpectedSprite)

End Function