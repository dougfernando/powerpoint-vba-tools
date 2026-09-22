Attribute VB_Name = "modText"
' @ManagedByDouglasTools
Option Explicit

Public Sub SetShapeDoNotAutofit(ByVal shp As Shape, ByRef changedCount As Long)
    Dim i As Long

    On Error GoTo Fail

    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            SetShapeDoNotAutofit shp.GroupItems(i), changedCount
        Next i
        Exit Sub
    End If

    If shp.HasTextFrame Then
        shp.TextFrame2.AutoSize = msoAutoSizeNone
        changedCount = changedCount + 1
    End If

    Exit Sub
Fail:
    ' Ignore shapes that do not support the requested operation.
End Sub

Public Sub ApplyStandardMargins(ByVal sr As ShapeRange, ByRef changedCount As Long)
    Dim shp As Shape
    Dim marginPoints As Single
    marginPoints = StandardMarginPoints()

    For Each shp In sr
        On Error Resume Next
        If shp.HasTextFrame Then
            With shp.TextFrame2
                .MarginBottom = marginPoints
                .MarginTop = marginPoints
                .MarginLeft = marginPoints
                .MarginRight = marginPoints
            End With
            changedCount = changedCount + 1
        End If
        On Error GoTo 0
    Next shp
End Sub

Public Sub CopyTextAndFrameFormat(ByVal sourceShape As Shape, ByVal targetShape As Shape)
    targetShape.TextFrame2.TextRange.Text = ""

    sourceShape.TextFrame2.TextRange.Copy
    targetShape.TextFrame2.TextRange.Paste

    With targetShape.TextFrame2
        .MarginLeft = sourceShape.TextFrame2.MarginLeft
        .MarginRight = sourceShape.TextFrame2.MarginRight
        .MarginTop = sourceShape.TextFrame2.MarginTop
        .MarginBottom = sourceShape.TextFrame2.MarginBottom
        .VerticalAnchor = sourceShape.TextFrame2.VerticalAnchor
        .WordWrap = sourceShape.TextFrame2.WordWrap
        .AutoSize = sourceShape.TextFrame2.AutoSize
    End With
End Sub

Public Sub SwapTextBetweenShapes(ByVal shape1 As Shape, ByVal shape2 As Shape)
    Dim tempText As String
    tempText = shape1.TextFrame2.TextRange.Text
    shape1.TextFrame2.TextRange.Text = shape2.TextFrame2.TextRange.Text
    shape2.TextFrame2.TextRange.Text = tempText
End Sub

Public Function IsAllowedFont(ByVal fontName As String) As Boolean
    Dim prefixes As Variant
    Dim i As Long

    prefixes = AllowedFontPrefixes()
    For i = LBound(prefixes) To UBound(prefixes)
        If Left$(fontName, Len(CStr(prefixes(i)))) = CStr(prefixes(i)) Then
            IsAllowedFont = True
            Exit Function
        End If
    Next i
End Function
