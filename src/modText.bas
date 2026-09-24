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

Public Sub CollapseRepeatedSpacesInShape(ByVal shp As Shape, ByRef changedCount As Long, ByRef removedCount As Long)
    Dim i As Long
    Dim cellShape As Shape
    Dim shapeRemovedCount As Long

    On Error GoTo Fail

    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            CollapseRepeatedSpacesInShape shp.GroupItems(i), changedCount, removedCount
        Next i
        Exit Sub
    End If

    If shp.HasTable = msoTrue Then
        Dim rowIndex As Long
        Dim columnIndex As Long

        For rowIndex = 1 To shp.Table.Rows.Count
            For columnIndex = 1 To shp.Table.Columns.Count
                Set cellShape = shp.Table.Cell(rowIndex, columnIndex).Shape
                shapeRemovedCount = shapeRemovedCount + CollapseRepeatedSpacesInTextRange(cellShape.TextFrame.TextRange)
            Next columnIndex
        Next rowIndex
    ElseIf ShapeHasText(shp) Then
        shapeRemovedCount = CollapseRepeatedSpacesInTextRange(shp.TextFrame.TextRange)
    End If

    If shapeRemovedCount > 0 Then
        changedCount = changedCount + 1
        removedCount = removedCount + shapeRemovedCount
    End If
    Exit Sub

Fail:
    ' Ignore shapes that do not expose editable text.
End Sub

Private Function CollapseRepeatedSpacesInTextRange(ByVal textRange As TextRange) As Long
    Dim content As String
    Dim i As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim runLength As Long
    Dim removed As Long
    Dim leftCharacter As String
    Dim rightCharacter As String

    content = textRange.Text
    i = Len(content)

    Do While i >= 1
        If Mid$(content, i, 1) = " " Then
            runEnd = i
            Do While i >= 1
                If Mid$(content, i, 1) <> " " Then Exit Do
                i = i - 1
            Loop
            runStart = i + 1
            runLength = runEnd - runStart + 1

            If runLength > 1 Then
                If runStart > 1 And runEnd < Len(content) Then
                    leftCharacter = Mid$(content, runStart - 1, 1)
                    rightCharacter = Mid$(content, runEnd + 1, 1)

                    If Not IsTextBoundaryCharacter(leftCharacter) And Not IsTextBoundaryCharacter(rightCharacter) Then
                        textRange.Characters(runStart + 1, runLength - 1).Delete
                        removed = removed + runLength - 1
                    End If
                End If
            End If
        Else
            i = i - 1
        End If
    Loop

    CollapseRepeatedSpacesInTextRange = removed
End Function

Private Function IsTextBoundaryCharacter(ByVal character As String) As Boolean
    IsTextBoundaryCharacter = (character = vbCr) Or _
                              (character = vbLf) Or _
                              (character = vbTab) Or _
                              (character = Chr$(11))
End Function

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
