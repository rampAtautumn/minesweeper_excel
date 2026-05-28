Attribute VB_Name = "mod_ui"

Option Explicit

'====================================================
' ASSET ROOTS
'====================================================

Public Function GetProjectRoot() As String

    GetProjectRoot = ThisWorkbook.Path

End Function

Public Function GetAssetsRoot() As String

    GetAssetsRoot = _
        GetProjectRoot() & _
        "\assets\sprites\"

End Function

'====================================================
' ASSET LOADER
'====================================================

Public Sub LoadAssets()

    On Error GoTo ErrorHandler

    AssetsRoot = GetAssetsRoot()

    If Len(Dir$(AssetsRoot, vbDirectory)) = 0 Then

        Err.Raise _
            vbObjectError + 1000, _
            "LoadAssets", _
            "Assets directory not found:" & vbCrLf & _
            AssetsRoot

    End If

    Set SpritePaths = _
        CreateObject("Scripting.Dictionary")

    RegisterGameplaySprites


    If Not VerifyAssets() Then

        Err.Raise _
            vbObjectError + 1001, _
            "LoadAssets", _
            "Asset verification failed."

    End If

    Exit Sub

ErrorHandler:

    MsgBox _
        "Asset loading failed:" & vbCrLf & _
        Err.Description, _
        vbCritical

    StopGameTimer

End Sub

'====================================================
' GAMEPLAY SPRITES
'====================================================

Private Sub RegisterGameplaySprites()

    RegisterSprite _
        "hidden", _
        "block.jpeg"

    RegisterSprite _
        "flag", _
        "flag.jpeg"

    RegisterSprite _
        "mine", _
        "mine.jpeg"

    RegisterSprite _
        "active_mine", _
        "active_mine.jpeg"

    RegisterSprite _
        "0", _
        "null.jpeg"

    RegisterSprite _
        "1", _
        "1.jpeg"

    RegisterSprite _
        "2", _
        "2.jpeg"

    RegisterSprite _
        "3", _
        "3.jpeg"

    RegisterSprite _
        "4", _
        "4.jpeg"

    RegisterSprite _
        "5", _
        "5.jpeg"

    RegisterSprite _
        "6", _
        "6.jpeg"

    RegisterSprite _
        "7", _
        "7.jpeg"

    RegisterSprite _
        "8", _
        "8.jpeg"
    RegisterSprite "background", "background.jpeg"

End Sub

'====================================================
' SPRITE REGISTRATION
'====================================================

Private Sub RegisterSprite( _
    ByVal SpriteKey As String, _
    ByVal FileName As String _
)

    Dim FullPath As String

    FullPath = AssetsRoot & FileName

    If SpritePaths.Exists(SpriteKey) Then

        Err.Raise _
            vbObjectError + 1002, _
            "RegisterSprite", _
            "Duplicate sprite key detected: " & _
            SpriteKey

    End If

    SpritePaths.Add _
        SpriteKey, _
        FullPath

End Sub

'====================================================
' SPRITE LOOKUP
'====================================================

Public Function GetSpritePath( _
    ByVal SpriteKey As String _
) As String

    If SpritePaths Is Nothing Then

        Err.Raise _
            vbObjectError + 1003, _
            "GetSpritePath", _
            "Sprite registry not initialized."

    End If

    If Not SpritePaths.Exists(SpriteKey) Then

        Err.Raise _
            vbObjectError + 1004, _
            "GetSpritePath", _
            "Sprite key not found: " & _
            SpriteKey

    End If

    GetSpritePath = _
        CStr(SpritePaths(SpriteKey))

End Function

'====================================================
' ASSET VERIFICATION
'====================================================

Public Function VerifyAssets() As Boolean

    Dim SpriteKey As Variant
    Dim AssetPath As String

    VerifyAssets = False

    If SpritePaths Is Nothing Then
        Exit Function
    End If

    If SpritePaths.Count = 0 Then
        Exit Function
    End If

    For Each SpriteKey In SpritePaths.Keys

        AssetPath = _
            CStr(SpritePaths(SpriteKey))

        If Not FileExists(AssetPath) Then

            MsgBox _
                "Missing asset file:" & vbCrLf & _
                AssetPath, _
                vbCritical

            Exit Function

        End If

        If Not IsValidImageExtension(AssetPath) Then

            MsgBox _
                "Invalid asset extension:" & vbCrLf & _
                AssetPath, _
                vbCritical

            Exit Function

        End If

    Next SpriteKey

    VerifyAssets = True

End Function

'====================================================
' FILE VALIDATION
'====================================================

Private Function FileExists( _
    ByVal FilePath As String _
) As Boolean

    On Error Resume Next

    FileExists = _
        (Len(Dir$(FilePath)) > 0)

    On Error GoTo 0

End Function

Private Function IsValidImageExtension( _
    ByVal FilePath As String _
) As Boolean

    Dim Extension As String

    Extension = _
        LCase$(Mid$( _
            FilePath, _
            InStrRev(FilePath, ".") + 1 _
        ))

    Select Case Extension

        Case "jpg", "jpeg", "png"

            IsValidImageExtension = True

        Case Else

            IsValidImageExtension = False

    End Select

End Function

'====================================================
' TILE SPRITE RESOLUTION
'====================================================

Public Function ResolveTileSprite( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long _
) As String

    If Not IsWithinBounds(RowIndex, ColIndex) Then

        ResolveTileSprite = "hidden"

        Exit Function

    End If

    '------------------------------
    ' Flagged tile
    '------------------------------

    If bandera(RowIndex, ColIndex) Then

        ResolveTileSprite = "flag"

        Exit Function

    End If

    '------------------------------
    ' Hidden tile
    '------------------------------

    If Not revelado(RowIndex, ColIndex) Then

        ResolveTileSprite = "hidden"

        Exit Function

    End If

    '------------------------------
    ' Mine tile
    '------------------------------

    If tablero(RowIndex, ColIndex) = -1 Then

        If RowIndex = ExplodedRow And _
           ColIndex = ExplodedCol Then

            ResolveTileSprite = _
                "active_mine"

        Else

            ResolveTileSprite = _
                "mine"

        End If

        Exit Function

    End If

    '------------------------------
    ' Number / empty tile
    '------------------------------

    ResolveTileSprite = _
        CStr(tablero(RowIndex, ColIndex))

End Function

'====================================================
' HUD DIGIT HELPERS
'====================================================

Public Function GetHudDigitSprite( _
    ByVal DigitValue As Long _
) As String

    If DigitValue < 0 Then
        DigitValue = 0
    End If

    If DigitValue > 9 Then
        DigitValue = 9
    End If

    GetHudDigitSprite = _
        "score_" & DigitValue

End Function

'====================================================
' REGISTRY UTILITIES
'====================================================

Public Function AssetRegistryInitialized() As Boolean

    AssetRegistryInitialized = _
        Not SpritePaths Is Nothing

End Function

Public Function AssetCount() As Long

    If SpritePaths Is Nothing Then

        AssetCount = 0

        Exit Function

    End If

    AssetCount = SpritePaths.Count

End Function

'====================================================
' DEBUG UTILITIES
'====================================================

Public Sub DebugPrintAssetRegistry()

    Dim SpriteKey As Variant

    If SpritePaths Is Nothing Then

        Debug.Print _
            "Sprite registry not initialized."

        Exit Sub

    End If

    Debug.Print _
        "===== ASSET REGISTRY ====="

    For Each SpriteKey In SpritePaths.Keys

        Debug.Print _
            SpriteKey & _
            " => " & _
            SpritePaths(SpriteKey)

    Next SpriteKey

End Sub

Public Sub DebugValidateAssets()

    If VerifyAssets() Then

        Debug.Print _
            "All assets validated successfully."

    Else

        Debug.Print _
            "Asset validation failed."

    End If

End Sub
