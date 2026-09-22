Attribute VB_Name = "modCommands_Text"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Text_DisableAutofit()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    Dim sld As Slide
    Dim shp As Shape
    Dim changedCount As Long

    Application.StartNewUndoEntry

    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            SetShapeDoNotAutofit shp, changedCount
        Next shp
    Next sld

    Notify "Ajuste automatico desativado em " & changedCount & " shapes de " & pres.Slides.Count & " slides."
    Exit Sub

ErrorHandler:
    Notify "Falha ao desativar o ajuste automatico: " & Err.Description
End Sub

Public Sub Cmd_Text_CopyIntoFilledShape()
    On Error GoTo ErrorHandler

    Dim shp1 As Shape, shp2 As Shape
    Dim textShape As Shape, targetShape As Shape

    If Not TryGetTwoSelectedShapes(shp1, shp2) Then
        Notify "Operacao nao executada: selecione exatamente dois shapes."
        Exit Sub
    End If

    If ShapeHasNoFill(shp1) And ShapeHasVisibleFill(shp2) Then
        Set textShape = shp1
        Set targetShape = shp2
    ElseIf ShapeHasNoFill(shp2) And ShapeHasVisibleFill(shp1) Then
        Set textShape = shp2
        Set targetShape = shp1
    Else
        Notify "Operacao nao executada: selecione um shape sem preenchimento e outro preenchido."
        Exit Sub
    End If

    If Not ShapeHasText(textShape) Then
        Notify "Operacao nao executada: o shape sem preenchimento nao possui texto."
        Exit Sub
    End If

    If Not ShapeSupportsText(targetShape) Then
        Notify "Operacao nao executada: o shape preenchido nao aceita texto."
        Exit Sub
    End If

    Application.StartNewUndoEntry
    CopyTextAndFrameFormat textShape, targetShape
    textShape.Delete

    Notify "Texto copiado para o shape preenchido; o shape de origem foi removido."
    Exit Sub

ErrorHandler:
    Notify "Falha ao copiar o texto: " & Err.Description
End Sub

Public Sub Cmd_Text_SetStandardMargins()
    On Error GoTo ErrorHandler

    Dim sr As ShapeRange
    If Not TryGetSelectedShapes(sr) Then
        Notify "Operacao nao executada: selecione um ou mais shapes."
        Exit Sub
    End If

    Dim changedCount As Long
    Application.StartNewUndoEntry
    ApplyStandardMargins sr, changedCount

    Notify "Margens de 0,2 cm aplicadas a " & changedCount & " shapes."
    Exit Sub

ErrorHandler:
    Notify "Falha ao aplicar as margens padrao: " & Err.Description
End Sub

Public Sub Cmd_Text_Swap()
    On Error GoTo ErrorHandler

    Dim shape1 As Shape, shape2 As Shape
    If Not TryGetTwoSelectedShapes(shape1, shape2) Then
        Notify "Operacao nao executada: selecione exatamente dois shapes."
        Exit Sub
    End If

    If Not ShapeSupportsText(shape1) Or Not ShapeSupportsText(shape2) Then
        Notify "Operacao nao executada: os dois shapes devem aceitar texto."
        Exit Sub
    End If

    Application.StartNewUndoEntry
    SwapTextBetweenShapes shape1, shape2

    Notify "Textos trocados entre os dois shapes selecionados."
    Exit Sub

ErrorHandler:
    Notify "Falha ao trocar os textos: " & Err.Description
End Sub

Public Sub Cmd_Text_RemoveManualLineBreaks()
    On Error GoTo ErrorHandler

    Dim shp As Shape
    If Not TryGetSingleSelectedShape(shp) Then
        Notify "Operacao nao executada: selecione exatamente um shape."
        Exit Sub
    End If

    If shp.HasTextFrame = msoFalse Then
        Notify "Operacao nao executada: o shape selecionado nao possui texto."
        Exit Sub
    End If

    If shp.TextFrame.HasText = msoFalse Then
        Notify "Operacao nao executada: o shape selecionado nao possui texto."
        Exit Sub
    End If

    Dim originalText As String
    Dim updatedText As String
    originalText = shp.TextFrame.TextRange.Text
    updatedText = originalText

    updatedText = Replace(updatedText, vbCrLf, " ")
    updatedText = Replace(updatedText, vbCr, " ")
    updatedText = Replace(updatedText, vbLf, " ")
    updatedText = Replace(updatedText, Chr$(11), " ")

    If updatedText = originalText Then
        Notify "Nenhuma quebra manual foi encontrada no shape selecionado."
        Exit Sub
    End If

    Application.StartNewUndoEntry
    shp.TextFrame.TextRange.Text = updatedText

    Notify "Quebras manuais removidas do shape selecionado."
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover as quebras manuais: " & Err.Description
End Sub
