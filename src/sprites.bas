Option Explicit

'=======================
' CONFIG LOCAL
'=======================

Private Const SPRITE_PREFIX As String = "Sprite_Duck_"
Private Const FRAME_COUNT As Long = 3

'=======================
' RUTAS (usa globales)
'=======================

Private Function BuildPath(subFolder As String, fileName As String) As String
    BuildPath = ThisWorkbook.Path & "\" & _
                ASSETS_ROOT & subFolder & fileName
End Function

Private Function GetDuckFramePath(frameNumber As Long) As String
    GetDuckFramePath = BuildPath(PATH_DUCKS, "duck_fly_" & frameNumber & ".png")
End Function

'=======================
' HELPERS
'=======================

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(SHEET_SPRITES)
    On Error GoTo 0
    
    If GetSheet Is Nothing Then
        Debug.Print "ERROR: Sheet not found -> " & SHEET_SPRITES
    End If
End Function

Private Function GetSpriteName(duckID As String) As String
    GetSpriteName = SPRITE_PREFIX & duckID
End Function

Private Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

'=======================
' CREATE
'=======================

Public Function CreateDuckSprite(duckID As String, startX As Double, startY As Double) As String
    Dim ws As Worksheet
    Dim shp As Shape
    Dim spriteName As String
    Dim imagePath As String
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    
    spriteName = GetSpriteName(duckID)
    imagePath = GetDuckFramePath(1)
    
    If Not FileExists(imagePath) Then
        Debug.Print "ERROR: Missing sprite -> " & imagePath
        Exit Function
    End If
    
    ' Eliminar si ya existe
    On Error Resume Next
    ws.Shapes(spriteName).Delete
    On Error GoTo 0
    
    ' Crear imagen
    Set shp = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=startX, _
        Top:=startY, _
        Width:=50, _
        Height:=50)
    
    If shp Is Nothing Then Exit Function
    
    shp.Name = spriteName
    CreateDuckSprite = spriteName
End Function

'=======================
' ANIMATION
'=======================

Public Sub SetDuckFrame(duckID As String, frameNumber As Long)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim spriteName As String
    Dim imagePath As String
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    
    spriteName = GetSpriteName(duckID)
    
    ' Loop de frames
    frameNumber = ((frameNumber - 1) Mod FRAME_COUNT) + 1
    
    On Error Resume Next
    Set shp = ws.Shapes(spriteName)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    imagePath = GetDuckFramePath(frameNumber)
    
    If Not FileExists(imagePath) Then Exit Sub
    
    shp.Fill.UserPicture imagePath
End Sub

'=======================
' MOVEMENT
'=======================

Public Sub MoveDuck(duckID As String, dx As Double, dy As Double)
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set shp = ws.Shapes(GetSpriteName(duckID))
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    shp.Left = shp.Left + dx
    shp.Top = shp.Top + dy
End Sub

'=======================
' REMOVE
'=======================

Public Sub RemoveDuck(duckID As String)
    Dim ws As Worksheet
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Shapes(GetSpriteName(duckID)).Delete
    On Error GoTo 0
End Sub

'=======================
' BOUNDS
'=======================

Public Function GetDuckBounds(duckID As String) As Variant
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    
    On Error Resume Next
    Set shp = ws.Shapes(GetSpriteName(duckID))
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Function
    
    GetDuckBounds = Array( _
        shp.Left, _
        shp.Top, _
        shp.Left + shp.Width, _
        shp.Top + shp.Height _
    )
End Function

'=======================
' HELPER ANIMATION
'=======================

Public Sub AnimateDuck(duckID As String, ByRef frameCounter As Long)
    frameCounter = frameCounter + 1
    SetDuckFrame duckID, frameCounter
End Sub