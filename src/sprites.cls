Option Explicit

'=======================
' CONFIG
'=======================

Private Const SHEET_NAME As String = "GameScreen"
Private Const SPRITE_PREFIX As String = "Sprite_Duck_"
Private Const FRAME_COUNT As Integer = 3 ' ajusta según tus PNGs

Private Property Get AssetsPath() As String
    AssetsPath = ThisWorkbook.Path & "\Assets\"
End Property

Private Function GetSheet() As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    
    If GetSheet Is Nothing Then
        MsgBox "Sheet '" & SHEET_NAME & "' not found!", vbCritical
    End If
End Function

Private Function GetSpriteName(duckID As String) As String
    GetSpriteName = SPRITE_PREFIX & duckID
End Function

'=======================
' FILE VALIDATION
'=======================

Private Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

'=======================
' SPRITE CREATION
'=======================

Public Function CreateDuckSprite(duckID As String, startX As Double, startY As Double) As String
    Dim ws As Worksheet
    Dim spriteName As String
    Dim shp As Shape
    Dim imagePath As String
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Function
    
    spriteName = GetSpriteName(duckID)
    imagePath = AssetsPath & "duck_fly_1.png"
    
    ' Validate file exists
    If Not FileExists(imagePath) Then
        MsgBox "Sprite file not found: " & imagePath, vbExclamation
        Exit Function
    End If
    
    ' Elimina si ya existe (evita duplicados)
    On Error Resume Next
    ws.Shapes(spriteName).Delete
    On Error GoTo 0
    
    ' Crear imagen directamente (MUCHO más eficiente que AddShape)
    On Error Resume Next
    Set shp = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=startX, _
        Top:=startY, _
        Width:=50, _
        Height:=50)
    On Error GoTo 0
    
    If shp Is Nothing Then
        MsgBox "Failed to create sprite for duck " & duckID, vbExclamation
        Exit Function
    End If
    
    shp.Name = spriteName
    CreateDuckSprite = spriteName
End Function

'=======================
' FRAME ANIMATION
'=======================

Public Sub SetDuckFrame(duckID As String, frameNumber As Integer)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim spriteName As String
    Dim imagePath As String
    
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub
    
    spriteName = GetSpriteName(duckID)
    
    ' Normalizar frame (loop automático)
    frameNumber = ((frameNumber - 1) Mod FRAME_COUNT) + 1
    
    On Error Resume Next
    Set shp = ws.Shapes(spriteName)
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    imagePath = AssetsPath & "duck_fly_" & frameNumber & ".png"
    
    ' Validate file exists before updating
    If Not FileExists(imagePath) Then
        ' Log and skip frame update
        Debug.Print "Warning: sprite file not found: " & imagePath
        Exit Sub
    End If
    
    On Error Resume Next
    shp.Fill.UserPicture imagePath
    On Error GoTo 0
End Sub

'=======================
' SPRITE POSITIONING
'=======================

Public Sub MoveDuck(duckID As String, dx As Double, dy As Double)
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = GetSheet().Shapes(GetSpriteName(duckID))
    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub
    
    shp.Left = shp.Left + dx
    shp.Top = shp.Top + dy
End Sub

Public Sub RemoveDuck(duckID As String)
    On Error Resume Next
    GetSheet().Shapes(GetSpriteName(duckID)).Delete
    On Error GoTo 0
End Sub

'=======================
' SPRITE BOUNDS
'=======================

Public Function GetDuckBounds(duckID As String) As Variant
    Dim shp As Shape
    
    On Error Resume Next
    Set shp = GetSheet().Shapes(GetSpriteName(duckID))
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
' ANIMATION HELPER
'=======================

Public Sub AnimateDuck(duckID As String, ByRef frameCounter As Integer)
    frameCounter = frameCounter + 1
    SetDuckFrame duckID, frameCounter
End Sub