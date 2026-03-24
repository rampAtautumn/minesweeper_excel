Attribute VB_Name = "Sprites"
Option Explicit

'=======================
' CONFIG
'=======================

Private Const SPRITE_PREFIX As String = "Sprite_Duck_"

'=======================
' PATH SYSTEM (PORTABLE)
'=======================

Private Function GetAssetsRoot() As String
    GetAssetsRoot = ThisWorkbook.Path & "\assets\sprites\"
End Function

Private Function BuildFolderPath(subFolder As String) As String
    BuildFolderPath = GetAssetsRoot() & subFolder
End Function

'=======================
' FILE SYSTEM (ROBUSTO)
'=======================

Private Function GetFirstImageFromFolder(subFolder As String) As String
    
    Dim folderPath As String
    folderPath = BuildFolderPath(subFolder)
    
    If Dir(folderPath, vbDirectory) = "" Then
        Debug.Print "❌ Carpeta no existe:", folderPath
        Exit Function
    End If
    
    Dim file As String
    file = Dir(folderPath & "*.png")
    
    If file <> "" Then
        GetFirstImageFromFolder = folderPath & file
        Debug.Print "✔ Imagen encontrada:", GetFirstImageFromFolder
    Else
        Debug.Print "❌ No hay PNG en:", folderPath
    End If

End Function

'=======================
' SHEET (FIX CRÍTICO)
'=======================

Private Function GetSheet() As Worksheet
    Set GetSheet = GameSheet   ' 🔥 MISMA HOJA QUE EL ENGINE
End Function

'=======================
' BACKGROUND
'=======================

Public Sub LoadBackground()

    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    Dim path As String
    path = GetFirstImageFromFolder(PATH_BACKGROUNDS)
    
    If path = "" Then Exit Sub

    On Error Resume Next
    ws.Shapes("Background").Delete
    On Error GoTo 0

    Dim shp As Shape
    
    Set shp = ws.Shapes.AddPicture(path, msoFalse, msoTrue, 0, 0, 800, 600)
    
    If shp Is Nothing Then Exit Sub
    
    shp.Name = "Background"
    shp.ZOrder msoSendToBack

End Sub

'=======================
' DUCK IMAGE
'=======================

Private Function GetDuckImage() As String
    GetDuckImage = GetFirstImageFromFolder(PATH_DUCKS)
End Function

'=======================
' CREATE DUCK
'=======================

Public Function CreateDuckSprite(duckID As String, x As Double, y As Double) As String

    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Function

    Dim path As String
    path = GetDuckImage()
    
    If path = "" Then
        Debug.Print "❌ No hay imágenes de pato"
        Exit Function
    End If

    Dim name As String
    name = SPRITE_PREFIX & duckID

    On Error Resume Next
    ws.Shapes(name).Delete
    On Error GoTo 0

    Dim shp As Shape
    
    Set shp = ws.Shapes.AddPicture(path, msoFalse, msoTrue, x, y, 60, 60)
    
    If shp Is Nothing Then Exit Function
    
    shp.Name = name

    CreateDuckSprite = name

End Function

'=======================
' MOVE
'=======================

Public Sub MoveDuck(duckID As String, dx As Double, dy As Double)

    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    Dim shp As Shape
    
    On Error Resume Next
    Set shp = ws.Shapes(SPRITE_PREFIX & duckID)
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
    ws.Shapes(SPRITE_PREFIX & duckID).Delete
    On Error GoTo 0

End Sub

'=======================
' CLEAR
'=======================

Public Sub ClearAllSprites()

    Dim ws As Worksheet
    Set ws = GetSheet()
    If ws Is Nothing Then Exit Sub

    Dim shp As Shape
    
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

End Sub

'=======================
' LOAD ALL
'=======================

Public Sub LoadAllAssets()

    Debug.Print "==== LOAD ASSETS ===="
    Debug.Print "ROOT:", GetAssetsRoot()
    
    LoadBackground
    
    Dim result As String
    result = CreateDuckSprite("test", 200, 200)
    
    If result = "" Then
        Debug.Print "❌ No se pudo crear el pato"
    Else
        Debug.Print "✔ Pato creado:", result
    End If

End Sub