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
    
    SpawnDuck
    LastSpawnTime = Timer
    
End Sub

Private Sub SpawnDuck()
    Dim d As Duck
    Dim duckID As String
    
    ' ID único simple
    duckID = CStr(DucksSpawned + 1)
    
    Set d = New Duck
    
    ' Spawn en posición aleatoria horizontal
    d.Init duckID, Rnd * 700, 300
    
    Ducks.Add d
    DucksSpawned = DucksSpawned + 1
    
End Sub

'=======================
' UPDATE + CLEANUP (CRÍTICO)
'=======================

Public Sub UpdateDucksSafe()
    Dim i As Long
    Dim d As Duck
    
    ' Recorrer al revés para eliminar sin romper colección
    For i = Ducks.Count To 1 Step -1
        
        Set d = Ducks(i)
        
        If d.Alive Then
            d.Update
        Else
            Ducks.Remove i
        End If
        
    Next i
End Sub

'=======================
' SHOOTING SYSTEM
'=======================

Public Sub HandleShot(ByVal x As Double, ByVal y As Double)
    Dim d As Duck
    
    ' Sin balas → no dispara
    If Bullets <= 0 Then Exit Sub
    
    Bullets = Bullets - 1
    
    For Each d In Ducks
        
        If d.Alive Then
            
            If d.IsHit(x, y) Then
                
                d.Kill
                
                Score = Score + 100
                DucksShot = DucksShot + 1
                
                Exit Sub ' solo mata uno por disparo
            End If
            
        End If
        
    Next d
End Sub

'=======================
' ROUND CONTROL
'=======================

Public Sub CheckRoundEnd()
    
    ' Aún faltan patos por aparecer
    If DucksSpawned < DucksPerRound Then Exit Sub
    
    ' Aún hay patos vivos
    If Ducks.Count > 0 Then Exit Sub
    
    ' Siguiente ronda
    CurrentRound = CurrentRound + 1
    
    If CurrentRound > MaxRound Then
        GameEnded = True
        GameRunning = False
        Exit Sub
    End If
    
    ' Reset de ronda
    DucksSpawned = 0
    DucksPerRound = DucksPerRound + 2
    
End Sub