Attribute VB_Name = "Engine"
Option Explicit

'=======================
' INITIALIZATION
'=======================

Public Sub InitializeGame()
    
    InitializeGlobals
    
    GameRunning = True
    
    UpdateUI   ' 👈 inicializar UI
    
    ScheduleNextFrame
    
End Sub

'=======================
' GAME LOOP
'=======================

Public Sub GameLoop()
    
    If Not GameRunning Then Exit Sub
    
    UpdateDeltaTime
    
    UpdateDuckSpawn
    UpdateDucksSafe
    
    CheckRoundEnd
    
    UpdateUI   ' 👈 actualizar score en pantalla
    
    ScheduleNextFrame
    
End Sub

'=======================
' FRAME SCHEDULING
'=======================

Private Sub ScheduleNextFrame()
    On Error Resume Next
    Application.OnTime Now + FrameDelay / 86400#, "GameLoop"
End Sub

'=======================
' TIME MANAGEMENT
'=======================

Private Sub UpdateDeltaTime()
    
    Dim currentTime As Double
    currentTime = Timer
    
    If currentTime < LastFrameTime Then
        LastFrameTime = currentTime
    End If
    
    DeltaTime = currentTime - LastFrameTime
    LastFrameTime = currentTime
    
End Sub

'=======================
' UI (SCORE)
'=======================

Private Sub UpdateUI()
    
    On Error Resume Next
    
    With ThisWorkbook.Sheets(SHEET_SPRITES)
        .Range("A1").Value = "Score: " & Score
        .Range("A2").Value = "Round: " & CurrentRound
        .Range("A3").Value = "Bullets: " & Bullets
    End With
    
    On Error GoTo 0
    
End Sub

'=======================
' GAME CONTROL
'=======================

Public Sub StartGame()
    InitializeGame
End Sub

Public Sub StopGame()
    GameRunning = False
End Sub

Public Sub RestartGame()
    StopGame
    InitializeGame
End Sub

'=======================
' INPUT - KEYBOARD
'=======================

Public Sub OnKeyPress(ByVal KeyCode As Integer)
    
    If KeyCode = vbKeySpace Then
        StartGame
    End If
    
End Sub

'=======================
' INPUT - MOUSE
'=======================

Public Sub UpdateMousePosition(ByVal target As Range)
    
    ' Centro de la celda (mejor precisión)
    MouseX = target.Left + target.Width / 2
    MouseY = target.Top + target.Height / 2
    
End Sub

Public Sub OnMouseClick()
    
    HandleShot MouseX, MouseY
    
End Sub

Public Sub HandleSheetClick(ByVal target As Range)
    
    UpdateMousePosition target
    OnMouseClick
    
End Sub