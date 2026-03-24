Option Explicit

'=======================
' GLOBAL STATE
'=======================

Public Sprites As Collection
Public GameRunning As Boolean

Public FrameDelay As Double
Public LastFrameTime As Double

'=======================
' INITIALIZATION
'=======================

Public Sub InitializeGame()
    Set Sprites = New Collection
    
    FrameDelay = 0.0333 ' ~30 FPS
    LastFrameTime = Timer
    
    LoadSprites
    
    GameRunning = True
    
    ScheduleNextFrame
End Sub

'=======================
' GAME LOOP (NON-BLOCKING)
'=======================

Public Sub GameLoop()
    If Not GameRunning Then Exit Sub
    
    UpdateDeltaTime
    UpdateSprites
    CheckCollisions
    RenderSprites
    
    ScheduleNextFrame
End Sub

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
    
    ' Manejo de overflow del Timer (medianoche)
    If currentTime < LastFrameTime Then
        LastFrameTime = currentTime
    End If
    
    LastFrameTime = currentTime
End Sub

'=======================
' SPRITES
'=======================

Public Sub LoadSprites()
    Dim sprite As Sprite
    
    ' Ejemplo: jugador
    Set sprite = New Sprite
    sprite.Load "hunter.png"
    
    Sprites.Add sprite
End Sub

'=======================
' UPDATE
'=======================

Public Sub UpdateSprites()
    Dim sprite As Sprite
    
    For Each sprite In Sprites
        sprite.Update
    Next sprite
End Sub

'=======================
' COLLISIONS (OPTIMIZED)
'=======================

Public Sub CheckCollisions()
    Dim i As Long, j As Long
    Dim spriteA As Sprite, spriteB As Sprite
    
    For i = 1 To Sprites.Count
        Set spriteA = Sprites(i)
        
        For j = i + 1 To Sprites.Count
            Set spriteB = Sprites(j)
            
            If Collide(spriteA, spriteB) Then
                HandleCollision spriteA, spriteB
            End If
        Next j
    Next i
End Sub

Public Function Collide(spriteA As Sprite, spriteB As Sprite) As Boolean
    Collide = Not ( _
        spriteA.Right < spriteB.Left Or _
        spriteA.Left > spriteB.Right Or _
        spriteA.Bottom < spriteB.Top Or _
        spriteA.Top > spriteB.Bottom _
    )
End Function

Public Sub HandleCollision(spriteA As Sprite, spriteB As Sprite)
    spriteA.HandleCollisionWith spriteB
    spriteB.HandleCollisionWith spriteA
End Sub

'=======================
' RENDER
'=======================

Public Sub RenderSprites()
    Dim sprite As Sprite
    
    For Each sprite In Sprites
        sprite.Render
    Next sprite
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
' INPUT (BÁSICO)
'=======================

Public Sub OnKeyPress(ByVal KeyCode As Integer)
    If KeyCode = vbKeySpace Then
        StartGame
    End If
End Sub