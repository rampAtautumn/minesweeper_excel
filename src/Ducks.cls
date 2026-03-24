Option Explicit

'=======================
' ID / STATE
'=======================

Private m_ID As String
Private m_Frame As Integer

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

'=======================
' INITIALIZATION (reemplaza constructor)
'=======================

Public Sub Init(ByVal duckID As String, ByVal startX As Double, ByVal startY As Double)
    m_ID = duckID
    
    m_X = startX
    m_Y = startY
    
    ' Velocidad inicial (puedes ajustar lógica)
    m_VelX = 2 + Rnd * 2
    m_VelY = -2 - Rnd * 2
    
    m_Frame = 1
    m_Alive = True
    
    ' Crear sprite visual
    CreateDuckSprite m_ID, m_X, m_Y
End Sub

Public Sub Update()
    If Not m_Alive Then Exit Sub
    
    ' Movimiento
    m_X = m_X + m_VelX
    m_Y = m_Y + m_VelY
    
    ' Rebote simple en bordes (ajusta límites)
    If m_X < 0 Or m_X > 800 Then m_VelX = -m_VelX
    If m_Y < 0 Or m_Y > 400 Then m_VelY = -m_VelY
    
    ' Animación
    m_Frame = m_Frame + 1
    SetDuckFrame m_ID, m_Frame
    
    ' Aplicar posición al sprite
    ApplyPosition
End Sub

Private Sub ApplyPosition()
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = ThisWorkbook.Sheets("GameScreen").Shapes("Sprite_Duck_" & m_ID)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    shp.Left = m_X
    shp.Top = m_Y
End Sub

Public Function IsHit(ByVal mouseX As Double, ByVal mouseY As Double) As Boolean
    Dim bounds As Variant
    
    bounds = GetDuckBounds(m_ID)
    If IsEmpty(bounds) Then Exit Function
    
    IsHit = ( _
        mouseX >= bounds(0) And mouseX <= bounds(2) And _
        mouseY >= bounds(1) And mouseY <= bounds(3) _
    )
End Function

Public Sub Kill()
    If Not m_Alive Then Exit Sub
    
    m_Alive = False
    RemoveDuck m_ID
End Sub

Public Property Get ID() As String
    ID = m_ID
End Property

Public Property Get Alive() As Boolean
    Alive = m_Alive
End Property

Dim d As Duck

Set d = New Duck
d.Init "1", 100, 200

' En tu loop:
d.Update

' Click:
If d.IsHit(MouseX, MouseY) Then
    d.Kill
End If

