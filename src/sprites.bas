Option Explicit

'=======================
' SPRITES - DUCKS
'=======================

Private Const SHEET_NAME As String = "GameScreen"

Private Property Get AssetsPath() As String
    AssetsPath = ThisWorkbook.Path & "\Assets\"
End Property

'---------------------------------
' Create a duck sprite on the sheet
'---------------------------------
Public Function CreateDuckSprite(duckID As String, startX As Double, startY As Double) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim spriteName As String
    Dim shp As Shape

    spriteName = "Sprite_Duck_" & duckID
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, startX, startY, 50, 50)

    With shp
        .Name = spriteName
        .LockAspectRatio = msoTrue
        .Line.Visible = msoFalse
        .Fill.Visible = msoTrue
        .Fill.UserPicture AssetsPath & "duck_fly_1.png"
    End With

    CreateDuckSprite = spriteName
End Function

'---------------------------------
' Change sprite frame (for animation)
'---------------------------------
Public Sub SetDuckFrame(duckID As String, frameNumber As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim shp As Shape
    Dim spriteName As String

    spriteName = "Sprite_Duck_" & duckID

    On Error Resume Next
    Set shp = ws.Shapes(spriteName)
    On Error GoTo 0

    If Not shp Is Nothing Then
        shp.Fill.UserPicture AssetsPath & "duck_fly_" & frameNumber & ".png"
    End If 
End Sub