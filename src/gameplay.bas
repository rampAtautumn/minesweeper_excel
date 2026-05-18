Attribute VB_Name = "mod_gameplay"

Option Explicit

'====================================================
' TILE REVEAL ENTRY
'====================================================

Public Sub RevealTile( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If revelado(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If bandera(RowIndex, ColIndex) Then
        Exit Sub
    End If

    revelado(RowIndex, ColIndex) = True

    MarkTileDirty RowIndex, ColIndex

    '------------------------------
    ' Mine hit
    '------------------------------

    If tablero(RowIndex, ColIndex) = -1 Then

        HandleGameOver RowIndex, ColIndex

        Exit Sub

    End If

    '------------------------------
    ' Flood fill
    '------------------------------

    If tablero(RowIndex, ColIndex) = 0 Then

        FloodReveal RowIndex, ColIndex

    End If

    '------------------------------
    ' Win validation
    '------------------------------

    If CheckVictoryCondition() Then

        HandleVictory

    End If

End Sub

'====================================================
' FLOOD FILL REVEAL
'====================================================

Public Sub FloodReveal( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim dr As Long
    Dim dc As Long

    Dim NextRow As Long
    Dim NextCol As Long

    For dr = -1 To 1

        For dc = -1 To 1

            If Not (dr = 0 And dc = 0) Then

                NextRow = RowIndex + dr
                NextCol = ColIndex + dc

                If IsWithinBounds(NextRow, NextCol) Then

                    If Not revelado(NextRow, NextCol) Then

                        If Not bandera(NextRow, NextCol) Then

                            revelado(NextRow, NextCol) = True

                            MarkTileDirty NextRow, NextCol

                            If tablero(NextRow, NextCol) = 0 Then

                                FloodReveal _
                                    NextRow, _
                                    NextCol

                            End If

                        End If

                    End If

                End If

            End If

        Next dc

    Next dr

End Sub

'====================================================
' FLAG TOGGLE
'====================================================

Public Sub ToggleFlag( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If revelado(RowIndex, ColIndex) Then
        Exit Sub
    End If

    bandera(RowIndex, ColIndex) = _
        Not bandera(RowIndex, ColIndex)

    If bandera(RowIndex, ColIndex) Then

        RemainingFlags = RemainingFlags - 1

    Else

        RemainingFlags = RemainingFlags + 1

    End If

    MarkTileDirty RowIndex, ColIndex

    UpdateMineCounterHUD

End Sub

'====================================================
' DOUBLE CLICK REVEAL
'====================================================

Public Sub DoubleReveal( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim RequiredFlags As Long
    Dim NearbyFlags As Long

    Dim dr As Long
    Dim dc As Long

    Dim CheckRow As Long
    Dim CheckCol As Long

    If GameOver Then
        Exit Sub
    End If

    If Not IsWithinBounds(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If Not revelado(RowIndex, ColIndex) Then
        Exit Sub
    End If

    If tablero(RowIndex, ColIndex) <= 0 Then
        Exit Sub
    End If

    RequiredFlags = tablero(RowIndex, ColIndex)

    NearbyFlags = CountAdjacentFlags( _
        RowIndex, _
        ColIndex _
    )

    If NearbyFlags <> RequiredFlags Then
        Exit Sub
    End If

    For dr = -1 To 1

        For dc = -1 To 1

            If Not (dr = 0 And dc = 0) Then

                CheckRow = RowIndex + dr
                CheckCol = ColIndex + dc

                If IsWithinBounds(CheckRow, CheckCol) Then

                    If Not bandera(CheckRow, CheckCol) Then

                        If Not revelado(CheckRow, CheckCol) Then

                            RevealTile _
                                CheckRow, _
                                CheckCol

                        End If

                    End If

                End If

            End If

        Next dc

    Next dr

End Sub

'====================================================
' FLAG COUNTING
'====================================================

Public Function CountAdjacentFlags( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Long

    Dim dr As Long
    Dim dc As Long

    Dim CheckRow As Long
    Dim CheckCol As Long

    CountAdjacentFlags = 0

    For dr = -1 To 1

        For dc = -1 To 1

            If Not (dr = 0 And dc = 0) Then

                CheckRow = RowIndex + dr
                CheckCol = ColIndex + dc

                If IsWithinBounds(CheckRow, CheckCol) Then

                    If bandera(CheckRow, CheckCol) Then

                        CountAdjacentFlags = _
                            CountAdjacentFlags + 1

                    End If

                End If

            End If

        Next dc

    Next dr

End Function

'====================================================
' VICTORY VALIDATION
'====================================================

Public Function CheckVictoryCondition() As Boolean

    Dim r As Long
    Dim c As Long

    CheckVictoryCondition = False

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) <> -1 Then

                If Not revelado(r, c) Then

                    Exit Function

                End If

            End If

        Next c

    Next r

    CheckVictoryCondition = True

End Function

'====================================================
' REVEAL SURROUNDING TILES
'====================================================

Public Sub RevealAdjacentTiles( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
)

    Dim dr As Long
    Dim dc As Long

    Dim NextRow As Long
    Dim NextCol As Long

    For dr = -1 To 1

        For dc = -1 To 1

            If Not (dr = 0 And dc = 0) Then

                NextRow = RowIndex + dr
                NextCol = ColIndex + dc

                If IsWithinBounds(NextRow, NextCol) Then

                    If Not bandera(NextRow, NextCol) Then

                        RevealTile _
                            NextRow, _
                            NextCol

                    End If

                End If

            End If

        Next dc

    Next dr

End Sub

'====================================================
' MASS DIRTY MARKING
'====================================================

Public Sub MarkAllVisibleTilesDirty()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If revelado(r, c) Then

                DirtyTiles(r, c) = True

            End If

        Next c

    Next r

End Sub

'====================================================
' GAMEPLAY DEBUG
'====================================================

Public Sub DebugPrintRevealState()

    Dim r As Long
    Dim c As Long

    Dim OutputLine As String

    Debug.Print "===== REVEAL STATE ====="

    For r = 1 To BoardRows

        OutputLine = vbNullString

        For c = 1 To BoardCols

            If revelado(r, c) Then

                OutputLine = OutputLine & " 1"

            Else

                OutputLine = OutputLine & " 0"

            End If

        Next c

        Debug.Print OutputLine

    Next r

End Sub

Public Sub DebugPrintFlagState()

    Dim r As Long
    Dim c As Long

    Dim OutputLine As String

    Debug.Print "===== FLAG STATE ====="

    For r = 1 To BoardRows

        OutputLine = vbNullString

        For c = 1 To BoardCols

            If bandera(r, c) Then

                OutputLine = OutputLine & " F"

            Else

                OutputLine = OutputLine & " ."

            End If

        Next c

        Debug.Print OutputLine

    Next r

End Sub