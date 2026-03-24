Option Explicit

'=======================
' ENTITY LOGIC - DUCKS
'=======================

' Update all ducks for the frame
Public Sub UpdateDucks()
    MoveDuck "duck1", 5, -3
    MoveDuck "duck2", 3, -2
End Sub

' Move a duck with speedX, speedY (screen wrapping)
Public Sub MoveDuck(name As String, speedX As Integer, speedY As Integer)
    With ActiveSheet.Shapes(name)
        .Left = .Left + speedX
        .Top = .Top + speedY

        ' If out of bounds, respawn
        If .Top < 0 Or .Left > 800 Then
            RespawnDuck name
        End If
    End With
End Sub

' Respawn a duck at the bottom with random X
Public Sub RespawnDuck(name As String)
    With ActiveSheet.Shapes(name)
        .Left = Int(Rnd() * 300) + 50
        .Top = 400
        .Visible = msoTrue
    End With
End Sub

' "Shoot" handler: hide duck shape
Public Sub ShootDuck()
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes(Application.Caller)
    shp.Visible = msoFalse
End Sub