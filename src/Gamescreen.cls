VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "GameScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    HandleSheetClick Target
End Sub