'Variables del estado del juego'
Public GameRunning as Boolean 
Public GamePaused as Boolean  
Public GameEnded as Boolean
Public CurrentRound as Integer 
Public const MaxRound as integer = 20

'Variables de puntuación y progreso'
Public Score as Long 
Public DucksShot as Integer 
Public DucksMissed as Integer 

'Variables del jugador/arma'
Public Bullets as Integer 
Public const MaxBullets as Integer = 3
Public PlayerShot As Boolean

'Posición del Mouse'
Public MouseX as Double
Public MouseY as Double

'Clase patos'
Public Ducks as Collection
Public DucksPerRound As Integer
Public DucksSpawned As Integer
Public SpawnDelay As Double
Public LastSpawnTime As Double

'Variables de velocidad'
Public GameSpeed as Double
Public const FrameDelay as Double = 0.0333

'Referencias a las hojas del juego'
Public Menusheet as Worksheet
Public GameSheet as Worksheet
Public PauseSheet as Worksheet
Public SpriteSheet as Worksheet

'Nombres de las hojas'
Public Const SHEET_GAME As String = "Game"
Public Const SHEET_MENU As String = "Menu"
Public Const SHEET_SPRITES As String = "Sprites"
Public Const SHEET_Pause As String = "Pause"