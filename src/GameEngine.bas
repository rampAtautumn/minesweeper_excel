Option Explicit

Private Sub StartGame()
    SetupEnviroment
    GameInit
    Gameloop
End Sub

Private Sub SetupEnviroment() 'Función para crear/verificar el entorno'
    HideGridlines

    Set Menusheet = GetOrCreateSheet(SHEET_MENU)
    Set GameSheet = GetOrCreateSheet(SHEET_GAME)
    set PauseSheet = GetOrCreateSheet(SHEET_Pause)
    set SpriteSheet = GetOrCreateSheet(SHEET_SPRITES)

End Sub


Private Function GetSheetIfExists(sheetName As String) As Worksheet 'Función para validar la existencia de una hoja'
    
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

Private Sub HideGridlines()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    
    ' Guardar la hoja actual
    Set originalSheet = ActiveSheet
    
    ' Iterar sobre todas las hojas
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate  ' Activar la hoja
        ActiveWindow.DisplayGridlines = False
    Next ws
    
    ' Volver a la hoja original
    originalSheet.Activate

End Sub

private Sub GameInit()
    'Estado del juego'

    GameRunning = True
    GamePaused = False
    GameEnded = False
    CurrentRound = 1

    'Estado del jugador'
    ResetVars

    'Iniciar collection'
    set ducks = new Collection
End Sub

Private Sub ResetVars() 'Sub para resetear variables del jugador' 
    
    Score = 0
    DucksShot = 0
    DucksMissed = 0
    Bullets = MaxBullets
    
    MouseX = ActiveWindow.Width / 2
    MouseY = ActiveWindow.Height / 2
    
    GameSpeed = 1

End Sub

Private Sub Gameloop()
    Do While GameRunning
        Frame
    Loop
End Sub

Private Sub Frame()

    Call InputProcess
    Call SpawnSystem
    Call UpdateEntities
    Call Collisioncheck
    Call StateUpdate
    Call RoundManager
    Call EventProcesser
    Call Wait(FrameDelay)

End Sub

private Sub InputProcess()
    MouseX = Application.CursorLeft
    MouseY = Application.CursorTop
End Sub

private Sub Collisioncheck()

    Dim duck As Variant

    For Each duck In Ducks
        
        'Aquí se verificará si el cursor golpea al pato
        
    Next duck

End Sub

private Sub StateUpdate()

 If DucksShot >= 10 Then
        CurrentRound = CurrentRound + 1
    End If

End Sub

private Sub EventProcesser()
    DoEvents  
End Sub

Private Sub Wait(Seconds As Double) ' función para evitar sobrecarga
    Dim EndTime As Double
    EndTime = Timer + Seconds
    Do While Timer < EndTime
        DoEvents  
    Loop
End Sub

Private Sub GameInit()

    GameRunning = True
    GamePaused = False
    GameEnded = False
    CurrentRound = 1

    ResetVars

    Set Ducks = New Collection

    InitRound

End Sub

Private Sub SpawnSystem()

    If DucksSpawned >= DucksPerRound Then Exit Sub
    
    If Timer - LastSpawnTime >= SpawnDelay Then
    
        SpawnDuck
        
        DucksSpawned = DucksSpawned + 1
        
        LastSpawnTime = Timer
        
    End If

End Sub

Private Sub RoundManager()

    If DucksShot + DucksMissed >= DucksPerRound Then
        
        CurrentRound = CurrentRound + 1
        
        InitRound
        
    End If

End Sub

private Function CreateDuck() As Object

    Dim duck As Object
    Set duck = CreateObject("Scripting.Dictionary")
    
    duck("x") = Rnd * 800
    duck("y") = Rnd * 400
    
    duck("vx") = 3
    duck("vy") = 0
    
    duck("alive") = True
    
    Set CreateDuck = duck

End Function

Sub SpawnDuck()

    Dim duck As Object
    
    Set duck = CreateDuck()
    
    Ducks.Add duck
    
End Sub

Sub UpdateEntities()

    Dim duck As Variant
    
    For Each duck In Ducks
    
        If duck("alive") Then
        
            duck("x") = duck("x") + duck("vx")
            duck("y") = duck("y") + duck("vy")
        
        End If
    
    Next duck

End Sub