Attribute VB_Name = "Engine"
Option Explicit

'=======================
' LOOP CONTROL
'=======================

Private TargetFrameTime As Double

'=======================
' PUBLIC ENTRY POINT
'=======================

Public Sub BootGame()

    StopGame

    InitializeGlobals
    SetupEnvironment
    LoadAssets
    InitializeGameState

    GameRunning = True
    GamePaused = False

    TargetFrameTime = FrameDelay

    RunGameLoop

End Sub

'=======================
' ENVIRONMENT SETUP
'=======================

Private Sub SetupEnvironment()

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Maximizar ventana
    Application.WindowState = xlMaximized

    GameSheet.Activate

    ' Limpiar
    GameSheet.Cells.Clear
    ClearAllShapes GameSheet

    ' Estilo "juego"
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80

    GameSheet.Cells.Interior.Color = RGB(0, 0, 0)

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

'=======================
' LOAD ASSETS
'=======================

Private Sub LoadAssets()

    ' Fondo + sprite de prueba
    LoadAllAssets

End Sub

'=======================
' INITIAL GAME STATE
'=======================

Private Sub InitializeGameState()

    CurrentRound = 1

    DucksSpawned = 0
    DucksShot = 0
    DucksMissed = 0

    Score = 0
    Bullets = MaxBullets

    LastFrameTime = Timer
    LastSpawnTime = Timer

End Sub

'=======================
' MAIN LOOP
'=======================

Private Sub RunGameLoop()

    Dim currentTime As Double
    Dim accumulator As Double
    Dim frameCounter As Long

    LastFrameTime = Timer
    accumulator = 0#
    frameCounter = 0

    Do While GameRunning

        currentTime = Timer

        ' DeltaTime robusto
        If currentTime < LastFrameTime Then
            DeltaTime = (86400# - LastFrameTime) + currentTime
        Else
            DeltaTime = currentTime - LastFrameTime
        End If

        LastFrameTime = currentTime
        accumulator = accumulator + DeltaTime

        If accumulator >= TargetFrameTime Then

            If Not GamePaused Then
                UpdateDuckSpawn
                UpdateDucksSafe
                CheckRoundEnd
                UpdateUI
            End If

            accumulator = accumulator - TargetFrameTime
            frameCounter = frameCounter + 1

        End If

        ' DoEvents only every 2 frames to avoid performance hit
        If frameCounter Mod 2 = 0 Then
            DoEvents
        End If

    Loop

End Sub

'=======================
' GAME LOGIC (FIX)
'=======================

Public Sub UpdateDuckSpawn()
    DuckManager.UpdateDuckSpawn
End Sub

Public Sub UpdateDucksSafe()
    DuckManager.UpdateDucksSafe
End Sub

Public Sub CheckRoundEnd()
    DuckManager.CheckRoundEnd
End Sub

'=======================
' UI
'=======================


Private Sub UpdateUI()

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SPRITES)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub

    ws.Range("A1").Value = "Score: " & Score
    ws.Range("A2").Value = "Round: " & CurrentRound
    ws.Range("A3").Value = "Bullets: " & Bullets

End Sub

'=======================
' CLEAR SHAPES
'=======================

Public Sub ClearAllShapes(ByVal ws As Worksheet)

    Dim shp As Shape

    On Error Resume Next

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    On Error GoTo 0

End Sub

'=======================
' GAME CONTROL
'=======================

Public Sub StartGame()
    BootGame
End Sub

Public Sub StopGame()
    GameRunning = False
End Sub

Public Sub PauseGame()
    GamePaused = True
End Sub

Public Sub ResumeGame()
    GamePaused = False
End Sub

Public Sub RestartGame()
    StopGame
    BootGame
End Sub

'=======================
' INPUT
'=======================

Public Sub OnKeyPress(ByVal KeyCode As Integer)

    Select Case KeyCode

        Case vbKeySpace
            StartGame

        Case vbKeyP
            If GamePaused Then ResumeGame Else PauseGame

    End Select

End Sub

Public Sub UpdateMousePosition(ByVal target As Range)

    MouseX = target.Left + target.Width / 2
    MouseY = target.Top + target.Height / 2

End Sub

Public Sub OnMouseClick()

    If Not GameRunning Then Exit Sub
    If GamePaused Then Exit Sub

    DuckManager.HandleShot MouseX, MouseY

End Sub

Public Sub HandleSheetClick(ByVal target As Range)

    UpdateMousePosition target
    OnMouseClick

End Sub