Option Explicit

'=======================
' ID / STATE
'=======================

Private m_ID As String
Private m_Frame As Long

'=======================
' POSITION / MOVEMENT
'=======================

Private m_X As Double
Private m_Y As Double

Private m_VelX As Double
Private m_VelY As Double

'=======================
' STATE
'=======================

Private m_Alive As Boolean

Private Const MAX_FRAME As Long = 3
Private Const SPRITE_SIZE As Double = 50

'=======================
' INITIALIZATION
'=======================

Public Sub Init(ByVal duckID As String, ByVal startX As Double, ByVal startY As Double)
    
    m_ID = duckID
    
    m_X = startX
    m_Y = startY
    
    Randomize
    
    m_VelX = 2 + Rnd * 2
    m_VelY = -2 - Rnd * 2
    
    m_Frame = 1
    m_Alive = True
    
    ' Crear sprite visual
    CreateDuckSprite m_ID, m_X, m_Y
    
End Sub

'=======================
' UPDATE
'=======================

Public Sub Update()
    
    If Not m_Alive Then Exit Sub
    
    ' Movimiento
    m_X = m_X + m_VelX
    m_Y = m_Y + m_VelY
    
    ' Rebote en bordes
    If m_X < 0 Or m_X > 800 Then m_VelX = -m_VelX
    If m_Y < 0 Or m_Y > 400 Then m_VelY = -m_VelY
    
    ' Animación controlada
    m_Frame = m_Frame + 1
    If m_Frame > MAX_FRAME Then m_Frame = 1
    
    SetDuckFrame m_ID, m_Frame
    
    ApplyPosition
    
End Sub

'=======================
' POSITION APPLY
'=======================

Private Sub ApplyPosition()
    Dim ws As Worksheet
    Dim shp As Shape
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SPRITES)
    Set shp = ws.Shapes("Sprite_Duck_" & m_ID)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    shp.Left = m_X
    shp.Top = m_Y
End Sub

'=======================
' HIT DETECTION
'=======================

Public Function IsHit(ByVal x As Double, ByVal y As Double) As Boolean
    
    Dim bounds As Variant
    
    bounds = GetDuckBounds(m_ID)
    If IsEmpty(bounds) Then Exit Function
    
    IsHit = ( _
        x >= bounds(0) And x <= bounds(2) And _
        y >= bounds(1) And y <= bounds(3) _
    )
    
End Function

'=======================
' STATE CHANGE
'=======================

Public Sub Kill()
    
    If Not m_Alive Then Exit Sub
    
    m_Alive = False
    
    RemoveDuck m_ID
    
End Sub

'=======================
' COLLISION PROPERTIES
'=======================

Public Property Get Left() As Double
    Left = m_X
End Property

Public Property Get Right() As Double
    Right = m_X + SPRITE_SIZE
End Property

Public Property Get Top() As Double
    Top = m_Y
End Property

Public Property Get Bottom() As Double
    Bottom = m_Y + SPRITE_SIZE
End Property

'=======================
' GETTERS
'=======================

Public Property Get ID() As String
    ID = m_ID
End Property

Public Property Get Alive() As Boolean
    Alive = m_Alive
End Property

Public Property Get X() As Double
    X = m_X
End Property

Public Property Get Y() As Double
    Y = m_Y
End Property

Public Property Get Frame() As Long
    Frame = m_Frame
End Property