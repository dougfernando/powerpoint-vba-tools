Attribute VB_Name = "modCommands_Shapes"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Shape_DeleteByFillColor()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
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

    Dim matchCount As Long
    matchCount = CountShapesWithFillColor(pres, selectedColor)

    If matchCount = 0 Then
        Notify "Nenhum shape com a mesma cor de preenchimento foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("delete-fill-color", _
        "Confirmacao necessaria para remover " & matchCount & " shapes pela cor.") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteShapesWithFillColor(pres, selectedColor)

    Notify deletedCount & " shapes com a cor selecionada foram removidos."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes pela cor: " & Err.Description
End Sub

Private Function CountShapesWithFillColor(ByVal pres As Presentation, ByVal rgbValue As Long) As Long
    Dim sld As Slide, shp As Shape
    Dim fillColor As Long
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If ShapeFillRgb(shp, fillColor) Then
                If fillColor = rgbValue Then CountShapesWithFillColor = CountShapesWithFillColor + 1
            End If
        Next shp
    Next sld
End Function

Private Function DeleteShapesWithFillColor(ByVal pres As Presentation, ByVal rgbValue As Long) As Long
    Dim sld As Slide, shp As Shape
    Dim i As Long, fillColor As Long
    For Each sld In pres.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If ShapeFillRgb(shp, fillColor) Then
                If fillColor = rgbValue Then
                    shp.Delete
                    DeleteShapesWithFillColor = DeleteShapesWithFillColor + 1
                End If
            End If
        Next i
    Next sld
End Function

Public Sub Cmd_Shape_DeleteSimilar()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    Dim shpRef As Shape
    If Not TryGetSingleSelectedShape(shpRef) Then
        Notify "Operacao nao executada: selecione exatamente um shape."
        Exit Sub
    End If

    Dim sldRef As Slide
    Set sldRef = ActiveWindow.View.Slide

    Dim matchCount As Long
    matchCount = CountSimilarShapes(pres, shpRef, sldRef)

    If matchCount = 0 Then
        Notify "Nenhum shape semelhante foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("delete-similar", _
        "Confirmacao necessaria para remover " & matchCount & " shapes semelhantes.") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteSimilarShapes(pres, shpRef, sldRef)

    Notify deletedCount & " shapes semelhantes foram removidos."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes semelhantes: " & Err.Description
End Sub

Private Function CountSimilarShapes(ByVal pres As Presentation, ByVal shpRef As Shape, ByVal sldRef As Slide) As Long
    Dim sld As Slide, shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If Not (sld.SlideID = sldRef.SlideID And shp.Id = shpRef.Id) Then
                If SameRect(shp, shpRef.Left, shpRef.Top, shpRef.Width, shpRef.Height, SHAPE_MATCH_TOLERANCE_PT) Then
                    CountSimilarShapes = CountSimilarShapes + 1
                End If
            End If
        Next shp
    Next sld
End Function

Private Function DeleteSimilarShapes(ByVal pres As Presentation, ByVal shpRef As Shape, ByVal sldRef As Slide) As Long
    Dim sld As Slide, shp As Shape
    Dim i As Long
    For Each sld In pres.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If Not (sld.SlideID = sldRef.SlideID And shp.Id = shpRef.Id) Then
                If SameRect(shp, shpRef.Left, shpRef.Top, shpRef.Width, shpRef.Height, SHAPE_MATCH_TOLERANCE_PT) Then
                    shp.Delete
                    DeleteSimilarShapes = DeleteSimilarShapes + 1
                End If
            End If
        Next i
    Next sld
End Function

Public Sub Cmd_Shape_RemoveOutsideSlide()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    Dim slideWidth As Single, slideHeight As Single
    slideWidth = SlideWidthPoints(pres)
    slideHeight = SlideHeightPoints(pres)

    Dim matchCount As Long
    matchCount = CountOutsideShapes(pres, slideWidth, slideHeight)

    If matchCount = 0 Then
        Notify "Nenhum shape completamente fora da area do slide foi encontrado."
        Exit Sub
    End If

    If Not ConfirmThroughStatus("remove-outside-slide", _
        "Confirmacao necessaria para remover " & matchCount & " shapes fora dos slides.") Then Exit Sub

    Application.StartNewUndoEntry

    Dim deletedCount As Long
    deletedCount = DeleteOutsideShapes(pres, slideWidth, slideHeight)

    Notify deletedCount & " shapes fora da area do slide foram removidos."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover shapes fora do slide: " & Err.Description
End Sub

Private Function CountOutsideShapes(ByVal pres As Presentation, ByVal slideWidth As Single, ByVal slideHeight As Single) As Long
    Dim sld As Slide, shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If ShapeIsCompletelyOutsideSlide(shp, slideWidth, slideHeight) Then CountOutsideShapes = CountOutsideShapes + 1
        Next shp
    Next sld
End Function

Private Function DeleteOutsideShapes(ByVal pres As Presentation, ByVal slideWidth As Single, ByVal slideHeight As Single) As Long
    Dim sld As Slide, shp As Shape
    Dim i As Long
    For Each sld In pres.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If ShapeIsCompletelyOutsideSlide(shp, slideWidth, slideHeight) Then
                shp.Delete
                DeleteOutsideShapes = DeleteOutsideShapes + 1
            End If
        Next i
    Next sld
End Function
