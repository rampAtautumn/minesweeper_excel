Attribute VB_Name = "mod_board"

Option Explicit

'====================================================
' BOARD INITIALIZATION
'====================================================

Public Sub InitializeEmptyBoard()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            tablero(r, c) = 0

            revelado(r, c) = False

            bandera(r, c) = False

            DirtyTiles(r, c) = True

        Next c

    Next r

End Sub

'====================================================
' MINE GENERATION
'====================================================

Public Sub GenerateMines()

    Dim MinesPlaced As Long

    Dim RandomRow As Long
    Dim RandomCol As Long

    MinesPlaced = 0

    Do While MinesPlaced < MineCount

        RandomRow = Int(BoardRows * Rnd) + 1
        RandomCol = Int(BoardCols * Rnd) + 1

        If tablero(RandomRow, RandomCol) <> -1 Then

            tablero(RandomRow, RandomCol) = -1

            MinesPlaced = MinesPlaced + 1

        End If

    Loop

End Sub

'====================================================
' ADJACENT COUNT GENERATION
'====================================================

Public Sub CalculateAdjacentCounts()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) <> -1 Then

                tablero(r, c) = CountAdjacentMines(r, c)

            End If

        Next c

    Next r

End Sub

'====================================================
' ADJACENT MINE COUNTER
'====================================================

Public Function CountAdjacentMines( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Long

    Dim dr As Long
    Dim dc As Long

    Dim CheckRow As Long
    Dim CheckCol As Long

    CountAdjacentMines = 0

    For dr = -1 To 1

        For dc = -1 To 1

            If Not (dr = 0 And dc = 0) Then

                CheckRow = RowIndex + dr
                CheckCol = ColIndex + dc

                If IsWithinBounds(CheckRow, CheckCol) Then

                    If tablero(CheckRow, CheckCol) = -1 Then

                        CountAdjacentMines = _
                            CountAdjacentMines + 1

                    End If

                End If

            End If

        Next dc

    Next dr

End Function

'====================================================
' BOARD RESET
'====================================================

Public Sub ClearBoardState()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            tablero(r, c) = 0

            revelado(r, c) = False

            bandera(r, c) = False

            DirtyTiles(r, c) = True

        Next c

    Next r

End Sub

'====================================================
' REVEAL ALL MINES
'====================================================

Public Sub RevealAllMines()

    Dim r As Long
    Dim c As Long

    For r = 1 To BoardRows

        For c = 1 To BoardCols

            If tablero(r, c) = -1 Then

                revelado(r, c) = True

                DirtyTiles(r, c) = True

            End If

        Next c

    Next r

End Sub

'====================================================
' BOARD DEBUG PRINT
'====================================================

Public Sub DebugPrintBoard()

    Dim r As Long
    Dim c As Long

    Dim OutputLine As String

    Debug.Print "===== BOARD ====="

    For r = 1 To BoardRows

        OutputLine = vbNullString

        For c = 1 To BoardCols

            OutputLine = OutputLine & _
                Format(tablero(r, c), " 0")

        Next c

        Debug.Print OutputLine

    Next r

End Sub

'====================================================
' SAFE TILE ACCESS
'====================================================

Public Function GetTileValue( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Integer

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        GetTileValue = 0

        Exit Function

    End If

    GetTileValue = tablero(RowIndex, ColIndex)

End Function

'====================================================
' TILE HELPERS
'====================================================

Public Function TileContainsMine( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Boolean

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        TileContainsMine = False

        Exit Function

    End If

    TileContainsMine = _
        (tablero(RowIndex, ColIndex) = -1)

End Function

Public Function TileIsRevealed( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Boolean

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        TileIsRevealed = False

        Exit Function

    End If

    TileIsRevealed = revelado(RowIndex, ColIndex)

End Function

Public Function TileIsFlagged( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As Boolean

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        TileIsFlagged = False

        Exit Function

    End If

    TileIsFlagged = bandera(RowIndex, ColIndex)

End Function