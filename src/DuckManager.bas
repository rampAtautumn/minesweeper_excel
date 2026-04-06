
Attribute VB_Name = "DuckManager"
Option Explicit

'=======================
' SPAWN CONTROL
'=======================

Public Sub UpdateDuckSpawn()
    
    ' No spawnear más de los permitidos en la ronda
    If DucksSpawned >= DucksPerRound Then Exit Sub
    
    ' Control de tiempo entre spawns
    If Timer - LastSpawnTime < SpawnDelay Then Exit Sub
    
    ' Validate spawn conditions before creating duck
    If DucksPerRound <= 0 Then Exit Sub
    
    SpawnDuck
    LastSpawnTime = Timer
    
End Sub

Private Sub SpawnDuck()
    
    Dim d As Duck
    Dim duckID As String
    Dim spawnX As Double
    Dim spawnY As Double
    
    ' ID único simple
    duckID = CStr(DucksSpawned + 1)
    
    ' Pre-calculate spawn position (avoid Rnd call during Add)
    spawnX = Rnd * 700
    spawnY = 300
    
    Set d = New Duck
    
    ' Initialize duck with pre-calculated position
    d.Init duckID, spawnX, spawnY
    
    ' Add to collection and increment counter atomically
    Ducks.Add d
    DucksSpawned = DucksSpawned + 1
    
End Sub

'=======================
' UPDATE + CLEANUP (CRÍTICO)
'=======================

Public Sub UpdateDucksSafe()
    
    Dim i As Long
    Dim d As Duck
    Dim duckCount As Long
    
    duckCount = Ducks.Count
    
    ' Guard against empty collection
    If duckCount = 0 Then Exit Sub
    
    ' Recorrer al revés para eliminar sin romper colección
    For i = duckCount To 1 Step -1
        
        ' Bounds validation
        If i <= Ducks.Count Then
            
            Set d = Ducks(i)
            
            If d.Alive Then
                d.Update
            Else
                Ducks.Remove i
            End If
            
        End If
        
    Next i
    
End Sub

'=======================
' SHOOTING SYSTEM
'=======================

Public Sub HandleShot(ByVal x As Double, ByVal y As Double)
    
    ' Sin balas → no dispara
    If Bullets <= 0 Then Exit Sub
    
    Bullets = Bullets - 1
    
    Dim d As Duck
    Dim i As Long
    
    ' Recorrer hacia atrás para detectar colisiones sin romper colección
    For i = Ducks.Count To 1 Step -1
        
        Set d = Ducks(i)
        
        If d.Alive Then
            
            If d.IsHit(x, y) Then
                
                d.Kill
                Score = Score + 100
                DucksShot = DucksShot + 1
                
                Exit Sub ' Solo mata uno por disparo
                
            End If
            
        End If
        
    Next i
    
End Sub

'=======================
' ROUND CONTROL
'=======================

Public Sub CheckRoundEnd()
    
    ' Aún faltan patos por aparecer
    If DucksSpawned < DucksPerRound Then Exit Sub
    
    ' Aún hay patos vivos (cache count)
    Dim duckCount As Long
    duckCount = Ducks.Count
    
    If duckCount > 0 Then Exit Sub
    
    ' Siguiente ronda
    CurrentRound = CurrentRound + 1
    
    ' Check if game ends
    If CurrentRound > MaxRound Then
        GameEnded = True
        GameRunning = False
        Exit Sub
    End If
    
    ' Reset de ronda
    DucksSpawned = 0
    DucksPerRound = DucksPerRound + 2
    
End Sub