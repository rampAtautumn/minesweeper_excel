Option Explicit

'=========================
' CONSTANTS - GAME BOUNDS
'=========================
Private Const GAME_WIDTH As Integer = 1000
Private Const GAME_HEIGHT As Integer = 600

'=========================
' UI CACHE (para evitar rewrites innecesarios)
'=========================
Private scoreCache As Long
Private bulletsCache As Integer
Private roundCache As Integer
Private missedCache As Integer

'=========================
' MAIN ENTRY POINT
'=========================

Public Sub StartGame()
    SetupEnvironment
    GameInit
    GameLoop
End Sub

'--------------------------
' INIT ENVIRONMENT & SHEETS
'--------------------------
Private Sub SetupEnvironment()
    ' ✅ OPTIMIZADO: No activar sheets, solo setear propiedad
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        ws.DisplayGridlines = False
    Next ws

    Set MenuSheet = GetOrCreateSheet(SHEET_MENU)
    Set GameSheet = GetOrCreateSheet(SHEET_GAME)
    Set PauseSheet = GetOrCreateSheet(SHEET_PAUSE)
    Set SpriteSheet = GetOrCreateSheet(SHEET_SPRITES)
End Sub

'--------------------------
' SHEET UTILS
'--------------------------
Private Function GetSheetIfExists(sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set GetSheetIfExists = ws
            Exit Function
        End If
    Next ws
    Set GetSheetIfExists = Nothing
End Function

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheetIfExists(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

'--------------------------
' GAME INITIALIZATION
'--------------------------
Private Sub GameInit()
    GameRunning = True
    GamePaused = False
    GameEnded = False
    CurrentRound = 1
    MaxBullets = 3
    Bullets = MaxBullets
    ReloadTime = 1.5
    LastShotTime = Timer
    LastFrameTime = Timer

    DucksPerRound = 10
    DucksSpawned = 0
    SpawnDelay = 1.2
    LastSpawnTime = Timer

    ResetVars

    Set Ducks = New Collection

    InitRound
    
    ' ✅ Inicializar caché
    scoreCache = 0
    bulletsCache = MaxBullets
    roundCache = 1
    missedCache = 0
End Sub

Private Sub ResetVars()
    Score = 0
    DucksShot = 0
    DucksMissed = 0
    Bullets = MaxBullets
    MouseX = Application.Width / 2
    MouseY = Application.Height / 2
    GameSpeed = 1
End Sub

'--------------------------
' MAIN GAME LOOP
'--------------------------
Private Sub GameLoop()
    Do While GameRunning
        Frame
    Loop
End Sub

Private Sub Frame()
    UpdateDeltaTime
    InputProcess
    SpawnSystem
    UpdateEntities
    CleanupDucks
    CollisionCheck
    ReloadSystem
    StateUpdate
    RoundManager
    CheckGameOver
    UpdateUI
    ' ✅ OPTIMIZADO: Eliminar DoEvents aquí (está en Wait)
    Wait FrameDelay
End Sub

'--------------------------
' FRAME SUBSYSTEMS
'--------------------------

Private Sub UpdateDeltaTime()
    Dim currentTime As Double
    currentTime = Timer
    DeltaTime = currentTime - LastFrameTime
    LastFrameTime = currentTime
End Sub

Private Sub InputProcess()
    MouseX = Application.CursorLeft
    MouseY = Application.CursorTop

    If GetAsyncKeyState(vbKeyLButton) <> 0 Then
        If Bullets > 0 Then
            PlayerShot = True
            Bullets = Bullets - 1
            LastShotTime = Timer
        Else
            PlayerShot = False
        End If
    Else
        PlayerShot = False
    End If
End Sub

Private Sub SpawnSystem()
    If DucksSpawned >= DucksPerRound Then Exit Sub
    If Timer - LastSpawnTime >= SpawnDelay Then
        SpawnDuck
        DucksSpawned = DucksSpawned + 1
        LastSpawnTime = Timer
    End If
End Sub

Private Sub SpawnDuck()
    Dim duck As Object
    Set duck = CreateDuck()
    Ducks.Add duck
End Sub

Private Function CreateDuck() As Object
    Dim duck As Object
    Set duck = CreateObject("Scripting.Dictionary")
    duck("x") = Rnd * 800
    duck("y") = Rnd * 400
    duck("vx") = 100
    duck("vy") = 0
    duck("alive") = True
    Set CreateDuck = duck
End Function

Private Sub UpdateEntities()
    Dim duck As Variant
    For Each duck In Ducks
        If duck("alive") Then
            duck("x") = duck("x") + duck("vx") * DeltaTime * 60
            duck("y") = duck("y") + duck("vy") * DeltaTime * 60
            ' ✅ OPTIMIZADO: Usar constante en lugar de hardcoded 1000
            If duck("x") > GAME_WIDTH Then
                duck("alive") = False
                DucksMissed = DucksMissed + 1
            End If
        End If
    Next duck
End Sub

Private Sub CleanupDucks()
    Dim i As Long, duck As Object
    For i = Ducks.Count To 1 Step -1
        Set duck = Ducks(i)
        If duck("alive") = False Then Ducks.Remove i
    Next i
End Sub

Private Sub CollisionCheck()
    Dim duck As Variant
    If Not PlayerShot Then Exit Sub
    
    ' ✅ OPTIMIZADO: Exit después de primer hit
    For Each duck In Ducks
        If duck("alive") Then
            If Abs(MouseX - duck("x")) < 30 And Abs(MouseY - duck("y")) < 30 Then
                duck("alive") = False
                DucksShot = DucksShot + 1
                Score = Score + 10
                Exit For  ' ✅ No chequear más patos
            End If
        End If
    Next duck
End Sub

Private Sub ReloadSystem()
    If Bullets = 0 And (Timer - LastShotTime >= ReloadTime) Then
        Bullets = MaxBullets
    End If
End Sub

Private Sub StateUpdate()
    If DucksShot >= 10 Then CurrentRound = CurrentRound + 1
End Sub

Private Sub RoundManager()
    If DucksShot + DucksMissed >= DucksPerRound Then
        CurrentRound = CurrentRound + 1
        InitRound
    End If
End Sub

Private Sub InitRound()
    DucksPerRound = Application.Max(10, CurrentRound * 2)
    DucksShot = 0
    DucksMissed = 0
    DucksSpawned = 0
    Set Ducks = New Collection
End Sub

Private Sub CheckGameOver()
    If CurrentRound > MaxRound Then
        GameRunning = False
        GameEnded = True
        MsgBox "¡Felicidades! Juego terminado. Tu puntaje: " & Score
    End If
End Sub

' ✅ OPTIMIZADO: Solo actualiza cells si valores cambiaron
Private Sub UpdateUI()
    If Score <> scoreCache Then
        GameSheet.Range("A1").Value = "Score: " & Score
        scoreCache = Score
    End If
    
    If Bullets <> bulletsCache Then
        GameSheet.Range("A2").Value = "Balas: " & Bullets
        bulletsCache = Bullets
    End If
    
    If CurrentRound <> roundCache Then
        GameSheet.Range("A3").Value = "Ronda: " & CurrentRound
        roundCache = CurrentRound
    End If
    
    If DucksMissed <> missedCache Then
        GameSheet.Range("A4").Value = "Fallados: " & DucksMissed
        missedCache = DucksMissed
    End If
End Sub

Private Sub Wait(seconds As Double)
    Dim EndTime As Double
    EndTime = Timer + seconds
    Do While Timer < EndTime
        DoEvents  ' ✅ Mantener aquí para responsividad
    Loop
End Sub