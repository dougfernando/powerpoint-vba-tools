Attribute VB_Name = "modCommands_Excel"
' @ManagedByDouglasTools
Option Explicit

Public Sub Cmd_Excel_ExportPresentation()
    On Error GoTo ErrorHandler

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    If Len(pres.Path) = 0 Then
        Notify "Operacao nao executada: salve a apresentacao antes de exportar."
        Exit Sub
    End If
    Dim succeeded As Boolean

    Dim xlApp As Object, xlWB As Object, xlWS As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWB = xlApp.Workbooks.Add
    Set xlWS = xlWB.Worksheets(1)

    WriteExportHeaders xlWS

    Dim row As Long
    row = 2

    Dim sld As Slide, shp As Shape
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            If ShapeHasText(shp) Then
                ExportShapeRow xlWS, row, sld, shp
                row = row + 1
            End If
        Next shp
    Next sld

    xlWS.Columns("A:I").AutoFit

    Dim outputPath As String
    outputPath = pres.Path & "\" & FileBaseName(pres.Name) & "_exported.xlsx"
    xlWB.SaveAs outputPath
    succeeded = True

CleanExit:
    On Error Resume Next
    If Not xlWB Is Nothing Then xlWB.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    On Error GoTo 0

    If succeeded Then
        Notify "Exportacao concluida: " & outputPath
    End If
    Exit Sub

ErrorHandler:
    Dim exportError As String
    exportError = Err.Description
    Notify "Falha ao exportar para Excel: " & exportError
    Resume CleanExit
End Sub

Private Sub WriteExportHeaders(ByVal ws As Object)
    ws.Cells(1, 1).Value = "Slide Index"
    ws.Cells(1, 2).Value = "Slide ID"
    ws.Cells(1, 3).Value = "Shape ID"
    ws.Cells(1, 4).Value = "Shape Name"
    ws.Cells(1, 5).Value = "Shape Text"
    ws.Cells(1, 6).Value = "X"
    ws.Cells(1, 7).Value = "Y"
    ws.Cells(1, 8).Value = "Width"
    ws.Cells(1, 9).Value = "Height"
End Sub

Private Sub ExportShapeRow(ByVal ws As Object, ByVal row As Long, ByVal sld As Slide, ByVal shp As Shape)
    ws.Cells(row, 1).Value = sld.SlideIndex
    ws.Cells(row, 2).Value = sld.SlideID
    ws.Cells(row, 3).Value = shp.Id
    ws.Cells(row, 4).Value = shp.Name
    ws.Cells(row, 6).Value = shp.Left
    ws.Cells(row, 7).Value = shp.Top
    ws.Cells(row, 8).Value = shp.Width
    ws.Cells(row, 9).Value = shp.Height

    Dim tr As TextRange
    Set tr = shp.TextFrame.TextRange
    ws.Cells(row, 5).Value = tr.Text
End Sub

Public Sub Cmd_Excel_UpdatePresentation()
    On Error GoTo ErrorHandler

    Dim succeeded As Boolean

    Dim pres As Presentation
    If Not TryGetTargetPresentation(pres) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        Exit Sub
    End If

    Dim filePath As String
    filePath = PickExcelFile()
    If Len(filePath) = 0 Then
        Notify "Operacao cancelada: nenhum arquivo Excel foi selecionado."
        Exit Sub
    End If

    Dim xlApp As Object, xlWB As Object, xlWS As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWB = xlApp.Workbooks.Open(filePath)
    Set xlWS = xlWB.Worksheets(1)

    Application.StartNewUndoEntry

    Dim updatedCount As Long
    updatedCount = ImportTextRows(pres, xlWS)
    succeeded = True

CleanExit:
    On Error Resume Next
    If Not xlWB Is Nothing Then xlWB.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlWS = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    On Error GoTo 0

    If succeeded Then
        Notify updatedCount & " shapes foram atualizados a partir do Excel."
    End If
    Exit Sub

ErrorHandler:
    Dim updateError As String
    updateError = Err.Description
    Notify "Falha ao atualizar pelo Excel: " & updateError
    Resume CleanExit
End Sub

Private Function ImportTextRows(ByVal pres As Presentation, ByVal ws As Object) As Long
    Dim row As Long
    row = 2

    Dim slideIndex As Long, shapeID As Long, newText As String
    Dim sld As Slide, shp As Shape

    Do While Not IsEmpty(ws.Cells(row, 1))
        slideIndex = CLng(ws.Cells(row, 1).Value)
        shapeID = CLng(ws.Cells(row, 3).Value)
        newText = CStr(ws.Cells(row, 5).Value)

        Set sld = Nothing
        Set shp = Nothing

        On Error Resume Next
        Set sld = pres.Slides(slideIndex)
        On Error GoTo 0

        If Not sld Is Nothing Then
            Set shp = FindShapeById(sld, shapeID)
            If Not shp Is Nothing Then
                If ShapeSupportsText(shp) Then
                    shp.TextFrame.TextRange.Text = newText
                    ImportTextRows = ImportTextRows + 1
                End If
            End If
        End If

        row = row + 1
    Loop
End Function

Private Function FindShapeById(ByVal sld As Slide, ByVal shapeID As Long) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If shp.Id = shapeID Then
            Set FindShapeById = shp
            Exit Function
        End If
    Next shp
End Function

Private Function PickExcelFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
End Function

Private Function FileBaseName(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        FileBaseName = Left$(fileName, dotPos - 1)
    Else
        FileBaseName = fileName
    End If
End Function
