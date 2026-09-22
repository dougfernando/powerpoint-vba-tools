Attribute VB_Name = "modSelection"
' @ManagedByDouglasTools
Option Explicit

Public Function TryGetSelectedShapes(ByRef result As ShapeRange, Optional ByVal exactCount As Long = 0) As Boolean
    On Error GoTo Fail

    If ActiveWindow Is Nothing Then
        Exit Function
    End If

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        Exit Function
    End If

    If exactCount > 0 Then
        If ActiveWindow.Selection.ShapeRange.Count <> exactCount Then
            Exit Function
        End If
    End If

    Set result = ActiveWindow.Selection.ShapeRange
    TryGetSelectedShapes = True
    Exit Function

Fail:
    TryGetSelectedShapes = False
End Function

Public Function TryGetSingleSelectedShape(ByRef shp As Shape) As Boolean
    Dim sr As ShapeRange
    If Not TryGetSelectedShapes(sr, 1) Then Exit Function
    Set shp = sr(1)
    TryGetSingleSelectedShape = True
End Function

Public Function TryGetTwoSelectedShapes(ByRef shp1 As Shape, ByRef shp2 As Shape) As Boolean
    Dim sr As ShapeRange
    If Not TryGetSelectedShapes(sr, 2) Then Exit Function
    Set shp1 = sr(1)
    Set shp2 = sr(2)
    TryGetTwoSelectedShapes = True
End Function
