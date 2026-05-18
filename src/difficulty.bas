Attribute VB_Name = "mod_difficulty"

Option Explicit

'====================================================
' DIFFICULTY ENUMERATION
'====================================================

Public Enum GameDifficulty

    DifficultyEasy = 1
    DifficultyMedium = 2
    DifficultyHard = 3
    DifficultyCustom = 4

End Enum

'====================================================
' CURRENT DIFFICULTY
'====================================================

Public CurrentDifficulty As GameDifficulty

'====================================================
' APPLY EASY
'====================================================

Public Sub SetEasyDifficulty()

    CurrentDifficulty = DifficultyEasy

    BoardRows = 8
    BoardCols = 8

    MineCount = 10

    RemainingFlags = MineCount

End Sub

'====================================================
' APPLY MEDIUM
'====================================================

Public Sub SetMediumDifficulty()

    CurrentDifficulty = DifficultyMedium

    BoardRows = 16
    BoardCols = 16

    MineCount = 40

    RemainingFlags = MineCount

End Sub

'====================================================
' APPLY HARD
'====================================================

Public Sub SetHardDifficulty()

    CurrentDifficulty = DifficultyHard

    BoardRows = 16
    BoardCols = 30

    MineCount = 99

    RemainingFlags = MineCount

End Sub

'====================================================
' CUSTOM DIFFICULTY
'====================================================

Public Sub SetCustomDifficulty( _
    ByVal RowsCount As Long, _
    ByVal ColsCount As Long, _
    ByVal MinesAmount As Long _
)

    ValidateCustomDifficulty _
        RowsCount, _
        ColsCount, _
        MinesAmount

    CurrentDifficulty = DifficultyCustom

    BoardRows = RowsCount
    BoardCols = ColsCount

    MineCount = MinesAmount

    RemainingFlags = MineCount

End Sub

'====================================================
' VALIDATION
'====================================================

Private Sub ValidateCustomDifficulty( _
    ByVal RowsCount As Long, _
    ByVal ColsCount As Long, _
    ByVal MinesAmount As Long _
)

    If RowsCount < 2 Then

        Err.Raise _
            vbObjectError + 4000, _
            "SetCustomDifficulty", _
            "Rows must be at least 2."

    End If

    If ColsCount < 2 Then

        Err.Raise _
            vbObjectError + 4001, _
            "SetCustomDifficulty", _
            "Columns must be at least 2."

    End If

    If MinesAmount <= 0 Then

        Err.Raise _
            vbObjectError + 4002, _
            "SetCustomDifficulty", _
            "Mine count must be greater than zero."

    End If

    If MinesAmount >= (RowsCount * ColsCount) Then

        Err.Raise _
            vbObjectError + 4003, _
            "SetCustomDifficulty", _
            "Mine count exceeds board capacity."

    End If

End Sub

'====================================================
' RESTART WITH DIFFICULTY
'====================================================

Public Sub RestartCurrentDifficulty()

    Select Case CurrentDifficulty

        Case DifficultyEasy

            SetEasyDifficulty

        Case DifficultyMedium

            SetMediumDifficulty

        Case DifficultyHard

            SetHardDifficulty

        Case Else

            Exit Sub

    End Select

    RestartEngine

End Sub

'====================================================
' DIFFICULTY NAME
'====================================================

Public Function GetDifficultyName() As String

    Select Case CurrentDifficulty

        Case DifficultyEasy

            GetDifficultyName = "Easy"

        Case DifficultyMedium

            GetDifficultyName = "Medium"

        Case DifficultyHard

            GetDifficultyName = "Hard"

        Case DifficultyCustom

            GetDifficultyName = "Custom"

        Case Else

            GetDifficultyName = "Unknown"

    End Select

End Function

'====================================================
' DIFFICULTY DEBUG
'====================================================

Public Sub DebugPrintDifficulty()

    Debug.Print _
        "===== DIFFICULTY ====="

    Debug.Print _
        "Mode: " & _
        GetDifficultyName()

    Debug.Print _
        "Rows: " & _
        BoardRows

    Debug.Print _
        "Columns: " & _
        BoardCols

    Debug.Print _
        "Mines: " & _
        MineCount

End Sub