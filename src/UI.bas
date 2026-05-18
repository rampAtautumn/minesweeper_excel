Attribute VB_Name = "mod_ui"

Option Explicit

'====================================================
' HUD CONFIGURATION
'====================================================

Public Const HUD_TOP_ROW As Long = 1
Public Const HUD_LEFT_COL As Long = 2

Public Const HUD_HEIGHT As Double = 32

Public Const HUD_DIGIT_WIDTH As Double = 24
Public Const HUD_DIGIT_HEIGHT As Double = 32

Public Const HUD_SPACING As Double = 2

'====================================================
' HUD INITIALIZATION
'====================================================

Public Sub InitializeHUD()

    Application.ScreenUpdating = False

    CreateHudBackground

    CreateMineCounter

    CreateTimerDisplay

    CreateRestartButton

    UpdateMineCounterHUD

    UpdateTimerHUD

    HudInitialized = True

    Application.ScreenUpdating = True

End Sub

'====================================================
' HUD BACKGROUND
'====================================================

Private Sub CreateHudBackground()

    Dim HudRange As Range

    SafeDeleteShape HUD_PREFIX & "background"

    Set HudRange = _
        GameSheet.Range( _
            GameSheet.Cells(1, 1), _
            GameSheet.Cells(3, BoardCols + 4) _
        )

    With GameSheet.Shapes.AddShape( _
        msoShapeRectangle, _
        HudRange.Left, _
        HudRange.Top, _
        HudRange.Width, _
        HudRange.Height _
    )

        .Name = HUD_PREFIX & "background"

        .Fill.ForeColor.RGB = RGB(45, 45, 45)

        .Line.Visible = msoFalse

        .Placement = xlMoveAndSize

    End With

End Sub

'====================================================
' MINE COUNTER CREATION
'====================================================

Private Sub CreateMineCounter()

    Dim i As Long

    Dim ShapeName As String

    Dim LeftPos As Double
    Dim TopPos As Double

    LeftPos = _
        GameSheet.Cells(2, 2).Left

    TopPos = _
        GameSheet.Cells(2, 2).Top

    For i = 1 To 3

        ShapeName = _
            HUD_PREFIX & _
            "mine_digit_" & i

        SafeDeleteShape ShapeName

        With GameSheet.Shapes.AddPicture( _
            Filename:=GetSpritePath("score_0"), _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=LeftPos, _
            Top:=TopPos, _
            Width:=HUD_DIGIT_WIDTH, _
            Height:=HUD_DIGIT_HEIGHT _
        )

            .Name = ShapeName

            .Placement = xlMoveAndSize

            .LockAspectRatio = msoFalse

        End With

        LeftPos = _
            LeftPos + HUD_DIGIT_WIDTH

    Next i

End Sub

'====================================================
' TIMER DISPLAY CREATION
'====================================================

Private Sub CreateTimerDisplay()

    Dim i As Long

    Dim ShapeName As String

    Dim LeftPos As Double
    Dim TopPos As Double

    LeftPos = _
        GameSheet.Cells(2, BoardCols - 1).Left

    TopPos = _
        GameSheet.Cells(2, BoardCols - 1).Top

    For i = 1 To 3

        ShapeName = _
            HUD_PREFIX & _
            "timer_digit_" & i

        SafeDeleteShape ShapeName

        With GameSheet.Shapes.AddPicture( _
            Filename:=GetSpritePath("score_0"), _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=LeftPos, _
            Top:=TopPos, _
            Width:=HUD_DIGIT_WIDTH, _
            Height:=HUD_DIGIT_HEIGHT _
        )

            .Name = ShapeName

            .Placement = xlMoveAndSize

            .LockAspectRatio = msoFalse

        End With

        LeftPos = _
            LeftPos + HUD_DIGIT_WIDTH

    Next i

End Sub

'====================================================
' RESTART BUTTON
'====================================================

Private Sub CreateRestartButton()

    Dim ButtonCell As Range

    Dim RestartButton As Shape

    SafeDeleteShape HUD_PREFIX & "restart"

    Set ButtonCell = _
        GameSheet.Cells(2, _
            Int(BoardCols / 2))

    Set RestartButton = _
        GameSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            ButtonCell.Left, _
            ButtonCell.Top, _
            60, _
            32 _
        )

    With RestartButton

        .Name = HUD_PREFIX & "restart"

        .TextFrame2.TextRange.Text = "RESET"

        .OnAction = "HandleRestartButton"

        .Fill.ForeColor.RGB = RGB(220, 220, 220)

        .Line.ForeColor.RGB = RGB(90, 90, 90)

        .Placement = xlMoveAndSize

    End With

End Sub

'====================================================
' MINE COUNTER UPDATE
'====================================================

Public Sub UpdateMineCounterHUD()

    UpdateHudNumber _
        RemainingFlags, _
        HUD_PREFIX & "mine_digit_"

End Sub

'====================================================
' TIMER UPDATE
'====================================================

Public Sub UpdateTimerHUD()

    UpdateHudNumber _
        CurrentElapsedSeconds, _
        HUD_PREFIX & "timer_digit_"

End Sub

'====================================================
' HUD NUMBER RENDERING
'====================================================

Private Sub UpdateHudNumber( _
    ByVal NumericValue As Long, _
    ByVal ShapePrefix As String _
)

    Dim DigitsText As String

    Dim i As Long

    Dim DigitValue As Long

    Dim ShapeName As String

    Dim ExistingShape As Shape

    NumericValue = Abs(NumericValue)

    If NumericValue > 999 Then
        NumericValue = 999
    End If

    DigitsText = _
        Right$("000" & CStr(NumericValue), 3)

    For i = 1 To 3

        DigitValue = _
            CLng(Mid$(DigitsText, i, 1))

        ShapeName = _
            ShapePrefix & i

        If ShapeExists(ShapeName) Then

            Set ExistingShape = _
                GameSheet.Shapes(ShapeName)

            ExistingShape.Fill.UserPicture _
                GetSpritePath( _
                    GetHudDigitSprite(DigitValue) _
                )

        End If

    Next i

End Sub

'====================================================
' HUD CLEANUP
'====================================================

Public Sub ClearHUD()

    Dim shp As Shape

    Dim ShapesToDelete As Collection

    Dim Item As Variant

    Set ShapesToDelete = New Collection

    For Each shp In GameSheet.Shapes

        If Left$(shp.Name, Len(HUD_PREFIX)) = _
            HUD_PREFIX Then

            ShapesToDelete.Add shp.Name

        End If

    Next shp

    For Each Item In ShapesToDelete

        SafeDeleteShape CStr(Item)

    Next Item

    HudInitialized = False

End Sub

'====================================================
' HUD REFRESH
'====================================================

Public Sub RefreshHUD()

    UpdateMineCounterHUD

    UpdateTimerHUD

End Sub

'====================================================
' SHAPE EXISTENCE CHECK
'====================================================

Private Function ShapeExists( _
    ByVal ShapeName As String _
) As Boolean

    On Error Resume Next

    ShapeExists = _
        Not GameSheet.Shapes(ShapeName) Is Nothing

    On Error GoTo 0

End Function

'====================================================
' BOARD/HUD ALIGNMENT
'====================================================

Public Sub RealignHUD()

    If Not HudInitialized Then
        Exit Sub
    End If

    ClearHUD

    InitializeHUD

End Sub

'====================================================
' UI DEBUG
'====================================================

Public Sub DebugPrintHUDState()

    Debug.Print _
        "===== HUD STATE ====="

    Debug.Print _
        "Remaining Flags: " & _
        RemainingFlags

    Debug.Print _
        "Elapsed Seconds: " & _
        CurrentElapsedSeconds

    Debug.Print _
        "HUD Initialized: " & _
        HudInitialized

End Sub