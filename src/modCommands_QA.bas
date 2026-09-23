Attribute VB_Name = "modCommands_QA"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_QA_ShowPlaceholderIDs()
    On Error GoTo ErrorHandler

    Dim sld As Slide
    If Not TryGetActiveSlide(sld) Then
        Notify "Operacao nao executada: nenhum slide ativo."
        Exit Sub
    End If

    Dim shp As Shape
    Dim msg As String

    For Each shp In sld.Shapes
        If shp.Type = msoPlaceholder Then
            msg = msg & "Name: " & shp.Name & _
                  " | Idx: " & shp.PlaceholderFormat.Idx & _
                  " | Type: " & shp.PlaceholderFormat.Type & vbCrLf
        End If
    Next shp

    If Len(msg) = 0 Then
        Notify "Nenhum placeholder foi encontrado no slide ativo."
    Else
        Notify "Placeholders: " & Replace(msg, vbCrLf, " | ")
    End If
    Exit Sub

ErrorHandler:
    Notify "Falha ao exibir os IDs dos placeholders: " & Err.Description
End Sub

Public Sub Cmd_QA_MarkNonAllowedFonts()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Application.StartNewUndoEntry

    Dim markedCount As Long
    Dim sld As Slide
    Dim shp As Shape
    For Each shp In targetShapes
        If ShapeUsesNonAllowedFont(shp) Then
            If TryGetShapeSlide(shp, sld) Then
                AddQAMarker sld, shp, QA_MARKER_FONT
                markedCount = markedCount + 1
            End If
        End If
    Next shp

    Notify markedCount & " shapes com fontes nao permitidas foram marcados no escopo " & CommandScopeLabel(scopeName) & "."
    Exit Sub

ErrorHandler:
    Notify "Falha ao verificar as fontes: " & Err.Description
End Sub

Private Function ShapeUsesNonAllowedFont(ByVal shp As Shape) As Boolean
    On Error GoTo Fail

    If Not ShapeHasText(shp) Then Exit Function

    Dim tr As TextRange
    Dim run As TextRange
    Set tr = shp.TextFrame.TextRange

    For Each run In tr.Runs
        If Not IsAllowedFont(run.Font.Name) Then
            ShapeUsesNonAllowedFont = True
            Exit Function
        End If
    Next run
    Exit Function
Fail:
    ShapeUsesNonAllowedFont = False
End Function

Private Sub AddQAMarker(ByVal sld As Slide, ByVal targetShape As Shape, ByVal markerType As String)
    Dim starSize As Single
    starSize = CSng(0.8 * CM_TO_POINTS)

    Dim centerX As Single, centerY As Single
    centerX = targetShape.Left + targetShape.Width / 2
    centerY = targetShape.Top + targetShape.Height / 2

    Dim marker As Shape
    Set marker = sld.Shapes.AddShape(msoShape5pointStar, centerX - starSize / 2, centerY - starSize / 2, starSize, starSize)
    marker.Fill.ForeColor.RGB = RGB(255, 0, 0)
    marker.Line.Visible = msoTrue
    marker.Line.ForeColor.RGB = RGB(255, 255, 0)
    marker.Line.Weight = 2
    marker.Tags.Add QA_MARKER_TAG_NAME, markerType
End Sub

Public Sub Cmd_QA_ClearMarkers()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Application.StartNewUndoEntry

    Dim shp As Shape
    Dim deletedCount As Long

    For Each shp In targetShapes
        If shp.Tags(QA_MARKER_TAG_NAME) <> "" Then
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next shp

    Notify deletedCount & " marcadores de QA foram removidos no escopo " & CommandScopeLabel(scopeName) & "."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover os marcadores de QA: " & Err.Description
End Sub

Public Sub Cmd_QA_CheckForbiddenClientNames()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Dim exceptionsInput As String
    exceptionsInput = InputBox("Exceptions separated by semicolons (;):", "Forbidden Client Names")

    Dim report As String
    report = BuildForbiddenClientReport(targetShapes, exceptionsInput)

    If Len(report) = 0 Then
        Notify "Nenhum nome proibido foi encontrado no escopo " & CommandScopeLabel(scopeName) & "."
    Else
        Notify "Ocorrencias no escopo " & CommandScopeLabel(scopeName) & ": " & Replace(report, vbCrLf, " | ")
    End If
    Exit Sub

ErrorHandler:
    Notify "Falha ao verificar nomes proibidos: " & Err.Description
End Sub

Private Function BuildForbiddenClientReport(ByVal targetShapes As Collection, ByVal exceptionsInput As String) As String
    Dim terms As Variant
    terms = ForbiddenClientTerms()

    Dim exceptions As Variant
    exceptions = Split(exceptionsInput, ";")

    Dim sld As Slide
    Dim shp As Shape
    Dim shapeText As String
    Dim i As Long
    Dim result As String

    For Each shp In targetShapes
        If ShapeHasText(shp) Then
            shapeText = shp.TextFrame2.TextRange.Text
            For i = LBound(terms) To UBound(terms)
                If Not IsExceptionTerm(CStr(terms(i)), exceptions) Then
                    If ContainsForbiddenTerm(shapeText, CStr(terms(i))) Then
                        If TryGetShapeSlide(shp, sld) Then
                            result = result & "Slide " & sld.SlideIndex & ": " & CStr(terms(i)) & vbCrLf
                        End If
                    End If
                End If
            Next i
        End If
    Next shp

    BuildForbiddenClientReport = result
End Function

Private Function IsExceptionTerm(ByVal term As String, ByVal exceptions As Variant) As Boolean
    Dim i As Long
    For i = LBound(exceptions) To UBound(exceptions)
        If StrComp(Trim$(term), Trim$(CStr(exceptions(i))), vbTextCompare) = 0 Then
            IsExceptionTerm = True
            Exit Function
        End If
    Next i
End Function

Private Function ContainsForbiddenTerm(ByVal haystack As String, ByVal needle As String) As Boolean
    If Len(Trim$(needle)) <= 2 Then
        ContainsForbiddenTerm = ContainsWholeWord(haystack, needle)
    Else
        ContainsForbiddenTerm = (InStr(1, haystack, needle, vbTextCompare) > 0)
    End If
End Function

Private Function ContainsWholeWord(ByVal haystack As String, ByVal needle As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(^|[^A-Za-z0-9À-ÿ])" & EscapeRegex(needle) & "([^A-Za-z0-9À-ÿ]|$)"
    re.IgnoreCase = True
    re.Global = False
    ContainsWholeWord = re.Test(haystack)
End Function

Private Function EscapeRegex(ByVal text As String) As String
    Dim chars As Variant, i As Long
    chars = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|", "/")
    EscapeRegex = text
    For i = LBound(chars) To UBound(chars)
        EscapeRegex = Replace(EscapeRegex, CStr(chars(i)), "\" & CStr(chars(i)))
    Next i
End Function
