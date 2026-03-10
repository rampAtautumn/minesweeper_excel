Option explicit

sub StartGame()
    GameRunning = true
    resetvar
End Sub

Sub ResetVars() 'Sub para resetear variables'
    score = 0
    CurrentRound = 0 'Actualizar en Gameloop a 1'
    DucksShot = 0
    DucksMissed = 0
    Bullets = MaxBullets
    MouseX = ActiveWindow.Width / 2
    MouseY = ActiveWindow.Height / 2
    GameSpeed = 1
End Sub

Private Function GetSheetIfExists(sheetName As String) As Worksheet
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set GetSheetIfExists = ws
            Exit Function
        End If
    Next ws
    
    Set GetSheetIfExists = Nothing
    
End Function

Pirvate Function GetOrCreateSheet(sheetName As String) As Worksheet
    
    Dim ws As Worksheet
    
    Set ws = GetSheetIfExists(sheetName)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet = ws

End Function