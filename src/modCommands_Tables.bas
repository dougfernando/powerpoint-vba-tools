Attribute VB_Name = "modCommands_Tables"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Table_ConvertToShapes()
    On Error GoTo ErrorHandler

    Dim tableShape As Shape
    If Not TryGetSingleSelectedShape(tableShape) Then
        Notify "Operacao nao executada: selecione exatamente uma tabela."
        Exit Sub
    End If

    If Not tableShape.HasTable Then
        Notify "Operacao nao executada: o shape selecionado nao e uma tabela."
        Exit Sub
    End If

    Dim sld As Slide
    If Not TryGetActiveSlide(sld) Then
        Notify "Operacao nao executada: nenhum slide ativo."
        Exit Sub
    End If

    Application.StartNewUndoEntry

    Dim tbl As Table
    Set tbl = tableShape.Table

    Dim rowNum As Long, colNum As Long
    Dim cellShape As Shape, newShape As Shape
    Dim createdCount As Long

    For rowNum = 1 To tbl.Rows.Count
        For colNum = 1 To tbl.Columns.Count
            Set cellShape = tbl.Cell(rowNum, colNum).Shape
            Set newShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, cellShape.Left, cellShape.Top, cellShape.Width, cellShape.Height)

            CopyCellToShape cellShape, newShape
            createdCount = createdCount + 1
        Next colNum
    Next rowNum

    Notify createdCount & " shapes foram criados a partir da tabela; a tabela original foi mantida."
    Exit Sub

ErrorHandler:
    Notify "Falha ao converter a tabela em shapes: " & Err.Description
End Sub

Private Sub CopyCellToShape(ByVal cellShape As Shape, ByVal targetShape As Shape)
    On Error Resume Next

    targetShape.TextFrame.TextRange.Text = cellShape.TextFrame.TextRange.Text

    With targetShape.TextFrame.TextRange.Font
        .Name = cellShape.TextFrame.TextRange.Font.Name
        .Size = cellShape.TextFrame.TextRange.Font.Size
        .Bold = cellShape.TextFrame.TextRange.Font.Bold
        .Italic = cellShape.TextFrame.TextRange.Font.Italic
        .Underline = cellShape.TextFrame.TextRange.Font.Underline
        .Color.RGB = cellShape.TextFrame.TextRange.Font.Color.RGB
    End With

    With targetShape.TextFrame2
        .MarginLeft = cellShape.TextFrame2.MarginLeft
        .MarginRight = cellShape.TextFrame2.MarginRight
        .MarginTop = cellShape.TextFrame2.MarginTop
        .MarginBottom = cellShape.TextFrame2.MarginBottom
        .VerticalAnchor = cellShape.TextFrame2.VerticalAnchor
        .WordWrap = cellShape.TextFrame2.WordWrap
    End With

    targetShape.Fill.Visible = cellShape.Fill.Visible
    targetShape.Fill.ForeColor.RGB = cellShape.Fill.ForeColor.RGB
    targetShape.Fill.Transparency = cellShape.Fill.Transparency
    targetShape.Line.Visible = msoFalse

    If cellShape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink Then
        targetShape.ActionSettings(ppMouseClick).Hyperlink.Address = cellShape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.Address
        targetShape.ActionSettings(ppMouseClick).Hyperlink.SubAddress = cellShape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress
    End If

    On Error GoTo 0
End Sub
