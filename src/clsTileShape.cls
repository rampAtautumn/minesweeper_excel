VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsTileShape"
Option Explicit

'====================================================
' PRIVATE MEMBERS
'====================================================

Private pRow As Long
Private pCol As Long

Private pShape As Shape

Private pCurrentSprite As String

Private pDirty As Boolean

'====================================================
' ROW
'====================================================

Public Property Get RowIndex() As Long

    RowIndex = pRow

End Property

Public Property Let RowIndex( _
    ByVal Value As Long _
)

    pRow = Value

End Property

'====================================================
' COLUMN
'====================================================

Public Property Get ColIndex() As Long

    ColIndex = pCol

End Property

Public Property Let ColIndex( _
    ByVal Value As Long _
)

    pCol = Value

End Property

'====================================================
' SHAPE REFERENCE
'====================================================

Public Property Get ShapeRef() As Shape

    Set ShapeRef = pShape

End Property

Public Property Set ShapeRef( _
    ByVal Value As Shape _
)

    Set pShape = Value

End Property

'====================================================
' CURRENT SPRITE
'====================================================

Public Property Get CurrentSprite() As String

    CurrentSprite = pCurrentSprite

End Property

Public Property Let CurrentSprite( _
    ByVal Value As String _
)

    pCurrentSprite = Value

End Property

'====================================================
' DIRTY STATE
'====================================================

Public Property Get Dirty() As Boolean

    Dirty = pDirty

End Property

Public Property Let Dirty( _
    ByVal Value As Boolean _
)

    pDirty = Value

End Property

'====================================================
' INITIALIZATION
'====================================================

Public Sub Initialize( _
    ByVal RowIndex As Long, _
    ByVal ColIndex As Long, _
    ByVal TileShape As Shape _
)

    pRow = RowIndex

    pCol = ColIndex

    Set pShape = TileShape

    pCurrentSprite = vbNullString

    pDirty = True

End Sub

'====================================================
' MARK DIRTY
'====================================================

Public Sub MarkDirty()

    pDirty = True

End Sub

'====================================================
' CLEAR DIRTY
'====================================================

Public Sub ClearDirty()

    pDirty = False

End Sub

'====================================================
' UPDATE SPRITE
'====================================================

Public Sub UpdateSprite( _
    ByVal SpriteKey As String _
)

    If pShape Is Nothing Then
        Exit Sub
    End If

    If pCurrentSprite = SpriteKey Then
        Exit Sub
    End If

    pShape.Fill.UserPicture _
        GetSpritePath(SpriteKey)

    pCurrentSprite = SpriteKey

    pDirty = False

End Sub

'====================================================
' SAFE DELETE
'====================================================

Public Sub DeleteShape()

    On Error Resume Next

    If Not pShape Is Nothing Then

        pShape.Delete

    End If

    Set pShape = Nothing

    On Error GoTo 0

End Sub

'====================================================
' REALIGN
'====================================================

Public Sub Realign()

    Dim TileCell As Range

    If pShape Is Nothing Then
        Exit Sub
    End If

    Set TileCell = _
        GetTileCell(pRow, pCol)

    With pShape

        .Left = TileCell.Left

        .Top = TileCell.Top

        .Width = TileSize

        .Height = TileSize

    End With

End Sub

'====================================================
' EXISTS
'====================================================

Public Function Exists() As Boolean

    Exists = _
        Not pShape Is Nothing

End Function

'====================================================
' DEBUG
'====================================================

Public Sub DebugPrint()

    Debug.Print _
        "Tile => (" & _
        pRow & "," & _
        pCol & ")" & _
        " Sprite=" & _
        pCurrentSprite & _
        " Dirty=" & _
        pDirty

End Sub