Attribute VB_Name = "modCommands_Shapes"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Shape_DeleteByFillColor()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()
    If scopeName = COMMAND_SCOPE_SELECTION Then
        Notify "Operacao nao executada: use o escopo slide ou apresentacao."
        Exit Sub
    End If

    Dim selectedShape As Shape
    If Not TryGetSingleSelectedShape(selectedShape) Then
        Notify "Operacao nao executada: selecione exatamente um shape."
        Exit Sub
    End If

    Dim selectedColor As Long
    If Not ShapeFillRgb(selectedShape, selectedColor) Then
        Notify "Operacao nao executada: o shape selecionado nao possui preenchimento visivel legivel."
        Exit Sub
    End If

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Dim matchCount As Long
    matchCount = CountShapesWithFillColor(targetShapes, selectedColor)

    If matchCount = 0 Then
        Notify "Nenhum shape com a mesma cor de preenchimento foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("delete-fill-color-" & scopeName, _
        "Confirme a remocao de " & matchCount & " shapes no escopo " & CommandScopeLabel(scopeName) & ".") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteShapesWithFillColor(targetShapes, selectedColor)

    Notify deletedCount & " shapes com a cor selecionada foram removidos no escopo " & CommandScopeLabel(scopeName) & "."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes pela cor: " & Err.Description
End Sub

Private Function CountShapesWithFillColor(ByVal targetShapes As Collection, ByVal rgbValue As Long) As Long
    Dim shp As Shape
    Dim fillColor As Long
    For Each shp In targetShapes
        If ShapeFillRgb(shp, fillColor) Then
            If fillColor = rgbValue Then CountShapesWithFillColor = CountShapesWithFillColor + 1
        End If
    Next shp
End Function

Private Function DeleteShapesWithFillColor(ByVal targetShapes As Collection, ByVal rgbValue As Long) As Long
    Dim shp As Shape
    Dim fillColor As Long
    For Each shp In targetShapes
        If ShapeFillRgb(shp, fillColor) Then
            If fillColor = rgbValue Then
                shp.Delete
                DeleteShapesWithFillColor = DeleteShapesWithFillColor + 1
            End If
        End If
    Next shp
End Function

Public Sub Cmd_Shape_DeleteSimilar()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()
    If scopeName = COMMAND_SCOPE_SELECTION Then
        Notify "Operacao nao executada: use o escopo slide ou apresentacao."
        Exit Sub
    End If

    Dim shpRef As Shape
    If Not TryGetSingleSelectedShape(shpRef) Then
        Notify "Operacao nao executada: selecione exatamente um shape."
        Exit Sub
    End If

    Dim sldRef As Slide
    If Not TryGetShapeSlide(shpRef, sldRef) Then
        Notify "Operacao nao executada: nao foi possivel identificar o slide da referencia."
        Exit Sub
    End If

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Dim matchCount As Long
    matchCount = CountSimilarShapes(targetShapes, shpRef, sldRef)

    If matchCount = 0 Then
        Notify "Nenhum shape semelhante foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("delete-similar-" & scopeName, _
        "Confirme a remocao de " & matchCount & " shapes no escopo " & CommandScopeLabel(scopeName) & ".") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteSimilarShapes(targetShapes, shpRef, sldRef)

    Notify deletedCount & " shapes semelhantes foram removidos no escopo " & CommandScopeLabel(scopeName) & "."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes semelhantes: " & Err.Description
End Sub

Private Function CountSimilarShapes(ByVal targetShapes As Collection, ByVal shpRef As Shape, ByVal sldRef As Slide) As Long
    Dim shp As Shape
    Dim sld As Slide
    For Each shp In targetShapes
        If TryGetShapeSlide(shp, sld) Then
            If Not (sld.SlideID = sldRef.SlideID And shp.Id = shpRef.Id) Then
                If SameRect(shp, shpRef.Left, shpRef.Top, shpRef.Width, shpRef.Height, SHAPE_MATCH_TOLERANCE_PT) Then
                    CountSimilarShapes = CountSimilarShapes + 1
                End If
            End If
        End If
    Next shp
End Function

Private Function DeleteSimilarShapes(ByVal targetShapes As Collection, ByVal shpRef As Shape, ByVal sldRef As Slide) As Long
    Dim shp As Shape
    Dim sld As Slide
    For Each shp In targetShapes
        If TryGetShapeSlide(shp, sld) Then
            If Not (sld.SlideID = sldRef.SlideID And shp.Id = shpRef.Id) Then
                If SameRect(shp, shpRef.Left, shpRef.Top, shpRef.Width, shpRef.Height, SHAPE_MATCH_TOLERANCE_PT) Then
                    shp.Delete
                    DeleteSimilarShapes = DeleteSimilarShapes + 1
                End If
            End If
        End If
    Next shp
End Function

Public Sub Cmd_Shape_RemoveOutsideSlide()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    Dim scopeName As String
    scopeName = CurrentCommandScope()

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Dim slideWidth As Single, slideHeight As Single
    slideWidth = SlideWidthPoints(pres)
    slideHeight = SlideHeightPoints(pres)

    Dim matchCount As Long
    matchCount = CountOutsideShapes(targetShapes, slideWidth, slideHeight)

    If matchCount = 0 Then
        Notify "Nenhum shape completamente fora da area do slide foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("remove-outside-slide-" & scopeName, _
        "Confirme a remocao de " & matchCount & " shapes no escopo " & CommandScopeLabel(scopeName) & ".") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteOutsideShapes(targetShapes, slideWidth, slideHeight)

    Notify deletedCount & " shapes fora da area do slide foram removidos no escopo " & CommandScopeLabel(scopeName) & "."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes fora do slide: " & Err.Description
End Sub

Private Function CountOutsideShapes(ByVal targetShapes As Collection, ByVal slideWidth As Single, ByVal slideHeight As Single) As Long
    Dim shp As Shape
    For Each shp In targetShapes
        If ShapeIsCompletelyOutsideSlide(shp, slideWidth, slideHeight) Then CountOutsideShapes = CountOutsideShapes + 1
    Next shp
End Function

Private Function DeleteOutsideShapes(ByVal targetShapes As Collection, ByVal slideWidth As Single, ByVal slideHeight As Single) As Long
    Dim shp As Shape
    For Each shp In targetShapes
        If ShapeIsCompletelyOutsideSlide(shp, slideWidth, slideHeight) Then
            shp.Delete
            DeleteOutsideShapes = DeleteOutsideShapes + 1
        End If
    Next shp
End Function
