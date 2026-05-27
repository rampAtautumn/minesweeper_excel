Attribute VB_Name = "mod_ui"

Option Explicit

'====================================================
' HUD CONSTANTS
'====================================================

Private Const HUD_MARGIN_X As Double = 8

Private Const HUD_PANEL_HEIGHT As Double = 42

Private Const RESET_BUTTON_SIZE As Double = 30

'====================================================
' HUD REFERENCES
'====================================================

Public RestartButtonShape As Shape

Public HudMineLabel As Shape
Public HudTimerLabel As Shape

'====================================================
' INITIALIZE HUD
'====================================================

Public Sub InitializeHUD()

    ClearHUD

    CreateHudBackground

    CreateMineLabel

    CreateRestartButton

    CreateFlagModeButton

    CreateTimerLabel

    HudInitialized = True

    RefreshHUD

End Sub

'====================================================
' HUD BACKGROUND
'====================================================

Private Sub CreateHudBackground()

    Dim BoardLeft As Double
    Dim BoardTop As Double
    Dim BoardWidth As Double

    SafeDeleteShape "hud_background"

    BoardLeft = GetBoardLeft()
    BoardTop = GetHudTop()
    BoardWidth = GetBoardWidth()

    With GameSheet.Shapes.AddShape( _
        msoShapeRectangle, _
        BoardLeft, _
        BoardTop, _
        BoardWidth, _
        HUD_PANEL_HEIGHT _
    )

        .Name = "hud_background"

        .Fill.ForeColor.RGB = RGB(192, 192, 192)

        .Line.ForeColor.RGB = RGB(110, 110, 110)

    End With

End Sub

'====================================================
' CREATE MINE LABEL
'====================================================

Private Sub CreateMineLabel()

    SafeDeleteShape "hud_mines"

    Set HudMineLabel = _
        GameSheet.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            GetBoardLeft() + 8, _
            GetHudTop() + 8, _
            90, _
            22 _
        )

    HudMineLabel.Name = "hud_mines"

    With HudMineLabel

        .Fill.Visible = msoFalse

        .Line.Visible = msoFalse

        .TextFrame2.TextRange.Font.Size = 12

        .TextFrame2.TextRange.Font.Bold = msoTrue

        .TextFrame2.TextRange.Text = _
            "Mines: " & RemainingFlags

    End With

End Sub

'====================================================
' CREATE TIMER LABEL
'====================================================

Private Sub CreateTimerLabel()

    SafeDeleteShape "hud_timer"

    Set HudTimerLabel = _
        GameSheet.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            GetBoardLeft() + GetBoardWidth() - 90, _
            GetHudTop() + 8, _
            80, _
            22 _
        )

    HudTimerLabel.Name = "hud_timer"

    With HudTimerLabel

        .Fill.Visible = msoFalse

        .Line.Visible = msoFalse

        .TextFrame2.TextRange.Font.Size = 12

        .TextFrame2.TextRange.Font.Bold = msoTrue

        .TextFrame2.TextRange.Text = "Time: 0"

    End With

End Sub

'====================================================
' CREATE RESTART BUTTON
'====================================================

Public Sub CreateRestartButton()

    Dim BtnLeft As Double
    Dim BtnTop As Double

    SafeDeleteShape "restart_button"

    BtnLeft = GetRestartButtonLeft()
    BtnTop = GetRestartButtonTop()

    Set RestartButtonShape = _
        GameSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            BtnLeft, _
            BtnTop, _
            RESET_BUTTON_SIZE, _
            RESET_BUTTON_SIZE _
        )

    RestartButtonShape.Name = _
        "restart_button"

    With RestartButtonShape

        .Fill.ForeColor.RGB = _
            RGB(235, 235, 235)

        .Line.ForeColor.RGB = _
            RGB(90, 90, 90)

        .Line.Weight = 1.5

        .TextFrame2.TextRange.Text = ":)"

        .TextFrame2.TextRange.Font.Size = 14

        .TextFrame2.TextRange.Font.Bold = msoTrue

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .OnAction = _
            "'" & ThisWorkbook.Name & "'!StartNewGame"

    End With

End Sub
Public Sub CreateFlagModeButton()

    Dim BtnLeft As Double
    Dim BtnTop As Double

    SafeDeleteShape "flag_mode_button"

    BtnLeft = _
        GetBoardLeft() + _
        GetBoardWidth() + 12

    BtnTop = _
        GetHudTop()

    Set FlagButtonShape = _
        GameSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            BtnLeft, _
            BtnTop, _
            70, _
            24 _
        )

    FlagButtonShape.Name = _
        "flag_mode_button"

    With FlagButtonShape

        .TextFrame2.TextRange.Text = _
            "FLAG: OFF"

        .Fill.ForeColor.RGB = _
            RGB(220, 220, 220)

        .Line.ForeColor.RGB = _
            RGB(90, 90, 90)

        .OnAction = _
            "ToggleFlagMode"

    End With

End Sub

'====================================================
' REFRESH HUD
'====================================================

Public Sub RefreshHUD()

    If Not HudInitialized Then
        Exit Sub
    End If

    If Not HudMineLabel Is Nothing Then

        HudMineLabel.TextFrame2.TextRange.Text = _
            "Mines: " & RemainingFlags

    End If

    If Not HudTimerLabel Is Nothing Then

        HudTimerLabel.TextFrame2.TextRange.Text = _
            "Time: " & CurrentElapsedSeconds

    End If

    UpdateRestartButtonState

End Sub

'====================================================
' UPDATE RESTART BUTTON
'====================================================

Public Sub UpdateRestartButtonState()

    If RestartButtonShape Is Nothing Then
        Exit Sub
    End If

    If GameWon Then

        RestartButtonShape.TextFrame2.TextRange.Text = ":D"

    ElseIf GameOver Then

        RestartButtonShape.TextFrame2.TextRange.Text = "X("

    Else

        RestartButtonShape.TextFrame2.TextRange.Text = ":)"

    End If

End Sub

'====================================================
' CLEAR HUD
'====================================================

Public Sub ClearHUD()

    SafeDeleteShape "hud_background"

    SafeDeleteShape "restart_button"

    SafeDeleteShape "hud_mines"

    SafeDeleteShape "hud_timer"

    SafeDeleteShape "flag_mode_button"

    ClearDifficultyButtons

    HudInitialized = False

    Set RestartButtonShape = Nothing

    Set HudMineLabel = Nothing

    Set HudTimerLabel = Nothing

End Sub

'====================================================
' REALIGN HUD
'====================================================

Public Sub RealignHUD()

    PositionRestartButton

    PositionLabels

End Sub

'====================================================
' POSITION LABELS
'====================================================

Private Sub PositionLabels()

    If Not HudMineLabel Is Nothing Then

        HudMineLabel.Left = _
            GetBoardLeft() + 8

        HudMineLabel.Top = _
            GetHudTop() + 8

    End If

    If Not HudTimerLabel Is Nothing Then

        HudTimerLabel.Left = _
            GetBoardLeft() + _
            GetBoardWidth() - 90

        HudTimerLabel.Top = _
            GetHudTop() + 8

    End If

End Sub

'====================================================
' POSITION RESTART BUTTON
'====================================================

Private Sub PositionRestartButton()

    If RestartButtonShape Is Nothing Then
        Exit Sub
    End If

    RestartButtonShape.Left = _
        GetRestartButtonLeft()

    RestartButtonShape.Top = _
        GetRestartButtonTop()

End Sub

'====================================================
' CREATE DIFFICULTY BUTTONS
'====================================================

Public Sub CreateDifficultyButtons()

    CreateDifficultyButton _
        "difficulty_easy", _
        "Easy", _
        1

    CreateDifficultyButton _
        "difficulty_medium", _
        "Medium", _
        2

    CreateDifficultyButton _
        "difficulty_hard", _
        "Hard", _
        3

End Sub

'====================================================
' CREATE SINGLE DIFFICULTY BUTTON
'====================================================

Private Sub CreateDifficultyButton( _
    ByVal ShapeName As String, _
    ByVal Caption As String, _
    ByVal Index As Long _
)

    Dim Btn As Shape

    Dim LeftPos As Double
    Dim TopPos As Double

    LeftPos = _
    GetBoardLeft() + _
    (GetBoardWidth() / 2) - 110 + _
    ((Index - 1) * 75)

    TopPos = _
        GetHudTop() - 34

    SafeDeleteShape ShapeName

    Set Btn = _
        GameSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            LeftPos, _
            TopPos, _
            60, _
            22 _
        )

    Btn.Name = ShapeName

    Btn.TextFrame2.TextRange.Text = Caption

    Btn.Fill.ForeColor.RGB = _
        RGB(220, 220, 220)

    Btn.Line.ForeColor.RGB = _
        RGB(90, 90, 90)

    Btn.OnAction = _
        "'" & ThisWorkbook.Name & "'!Set" & Caption & "AndRestart"

End Sub
Public Sub UpdateFlagModeButton()

    If FlagButtonShape Is Nothing Then
        Exit Sub
    End If

    If FlagModeEnabled Then

        FlagButtonShape.TextFrame2.TextRange.Text = _
            "FLAG: ON"

        FlagButtonShape.Fill.ForeColor.RGB = _
            RGB(255, 220, 120)

    Else

        FlagButtonShape.TextFrame2.TextRange.Text = _
            "FLAG: OFF"

        FlagButtonShape.Fill.ForeColor.RGB = _
            RGB(220, 220, 220)

    End If

End Sub
'====================================================
' CLEAR DIFFICULTY BUTTONS
'====================================================

Public Sub ClearDifficultyButtons()

    SafeDeleteShape "difficulty_easy"

    SafeDeleteShape "difficulty_medium"

    SafeDeleteShape "difficulty_hard"

End Sub

'====================================================
' DIFFICULTY ACTIONS
'====================================================

Public Sub SetEasyAndRestart()

    SetEasyDifficulty

    StartNewGame

End Sub

Public Sub SetMediumAndRestart()

    SetMediumDifficulty

    StartNewGame

End Sub

Public Sub SetHardAndRestart()

    SetHardDifficulty

    StartNewGame

End Sub

'====================================================
' HELPERS
'====================================================

Private Function GetBoardLeft() As Double

    GetBoardLeft = _
        GameSheet.Cells( _
            BoardOriginRow, _
            BoardOriginCol _
        ).Left

End Function

Private Function GetBoardTop() As Double

    GetBoardTop = _
        GameSheet.Cells( _
            BoardOriginRow, _
            BoardOriginCol _
        ).Top

End Function

Private Function GetBoardWidth() As Double

    GetBoardWidth = _
        BoardCols * TileSize

End Function

Private Function GetHudTop() As Double

    GetHudTop = _
        GetBoardTop() - _
        HUD_PANEL_HEIGHT - 8

End Function

Private Function GetRestartButtonLeft() As Double

    GetRestartButtonLeft = _
        GetBoardLeft() + _
        (GetBoardWidth() / 2) - _
        (RESET_BUTTON_SIZE / 2)

End Function

Private Function GetRestartButtonTop() As Double

    GetRestartButtonTop = _
        GetHudTop() + 4

End Function
Public Sub ToggleFlagMode()

    FlagModeEnabled = _
        Not FlagModeEnabled

    UpdateFlagModeButton

End Sub