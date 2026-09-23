Attribute VB_Name = "modCommands_Text"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Text_DisableAutofit()
    On Error GoTo ErrorHandler

    Dim scopeName As String
    scopeName = CurrentCommandScope()

    Dim targetShapes As Collection
    If Not TryGetShapesForScope(scopeName, targetShapes) Then
        Notify "Operacao nao executada: o escopo " & CommandScopeLabel(scopeName) & " nao esta disponivel."
        Exit Sub
    End If

    Dim shp As Shape
    Dim changedCount As Long

    Application.StartNewUndoEntry

    For Each shp In targetShapes
        SetShapeDoNotAutofit shp, changedCount
    Next shp

    Notify "Ajuste automatico desativado em " & changedCount & " shapes no escopo " & CommandScopeLabel(scopeName) & "."
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

    Dim sr As ShapeRange
    If Not TryGetSelectedShapes(sr) Then
        Notify "Operacao nao executada: selecione um ou mais shapes."
        Exit Sub
    End If

    Dim shp As Shape
    Dim originalText As String
    Dim updatedText As String
    Dim changedCount As Long

    For Each shp In sr
        If ShapeHasText(shp) Then
            originalText = shp.TextFrame.TextRange.Text
            updatedText = originalText

            updatedText = Replace(updatedText, vbCrLf, " ")
            updatedText = Replace(updatedText, vbCr, " ")
            updatedText = Replace(updatedText, vbLf, " ")
            updatedText = Replace(updatedText, Chr$(11), " ")

            If updatedText <> originalText Then
                If changedCount = 0 Then Application.StartNewUndoEntry
                shp.TextFrame.TextRange.Text = updatedText
                changedCount = changedCount + 1
            End If
        End If
    Next shp

    If changedCount = 0 Then
        Notify "Nenhuma quebra manual foi encontrada nos shapes selecionados."
    Else
        Notify "Quebras manuais removidas de " & changedCount & " shapes selecionados."
    End If
    Exit Sub

ErrorHandler:
    Notify "Falha ao remover as quebras manuais: " & Err.Description
End Sub
