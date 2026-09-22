Attribute VB_Name = "modMacroLauncher"
Option Explicit
' @ManagedByDouglasTools

Private launcher As frmMacroLauncher
Private running As Boolean

Public Sub ShowMacroLauncher()
    On Error GoTo Failed
    If launcher Is Nothing Then Set launcher = New frmMacroLauncher
    launcher.PrepareForDisplay
    launcher.Show vbModeless
    Exit Sub
Failed:
    Notify "Nao foi possivel abrir as macros: " & Err.Description
End Sub

Public Sub RibbonShowMacroLauncher(ByVal control As Office.IRibbonControl)
    ShowMacroLauncher
End Sub

Public Sub CloseMacroLauncher()
    On Error Resume Next
    CancelNotificationClear
    CancelPendingConfirmation
    If Not launcher Is Nothing Then Unload launcher
    Set launcher = Nothing
    On Error GoTo 0
End Sub

Public Sub ClearLauncherResult()
    On Error Resume Next
    If Not launcher Is Nothing Then launcher.ClearResult
    On Error GoTo 0
End Sub

Public Sub ShowLauncherResult(ByVal message As String, ByVal state As String)
    On Error Resume Next
    If launcher Is Nothing Then Set launcher = New frmMacroLauncher
    launcher.Show vbModeless
    launcher.DisplayResult message, state
    ScheduleNotificationClear
    On Error GoTo 0
End Sub

Public Sub Auto_Close()
    CloseMacroLauncher
End Sub

Public Function ToolsHealthCheck() As String
    Dim probe As New frmMacroLauncher
    Dim entries As Collection
    Set entries = CommandCatalog()
    Load probe
    If probe.lstMacro.ListCount <> entries.Count Then Err.Raise 5, , "Catalog mismatch"
    Unload probe
    ToolsHealthCheck = "DouglasPowerPointTools:OK"
End Function

Public Sub PopulateCategories(ByVal cbo As Object)
    Dim item As Variant, categories As New Collection, category As Variant
    cbo.Clear
    cbo.AddItem "Todas"
    For Each item In CommandCatalog()
        On Error Resume Next
        categories.Add CStr(item(1)), CStr(item(1))
        On Error GoTo 0
    Next item
    For Each category In categories
        cbo.AddItem CStr(category)
    Next category
    cbo.ListIndex = 0
End Sub

Public Sub PopulateCommands(ByVal cbo As Object, ByVal category As String)
    Dim item As Variant
    cbo.Clear
    cbo.ColumnCount = 3
    cbo.ColumnWidths = "360 pt;0 pt;0 pt"
    For Each item In CommandCatalog()
        If category = "Todas" Or category = CStr(item(1)) Then
            cbo.AddItem CStr(item(2)) & " (" & CStr(item(1)) & ")"
            cbo.List(cbo.ListCount - 1, 1) = CStr(item(0))
            cbo.List(cbo.ListCount - 1, 2) = CStr(item(3))
        End If
    Next item
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
End Sub

Public Function MacroLauncherRunSelected(ByVal cbo As Object) As String
    Dim target As Presentation
    Dim errorMessage As String

    ClearNotification
    If running Then
        Notify "Aguarde a operacao atual terminar."
        GoTo Done
    End If
    If cbo.ListIndex < 0 Then
        Notify "Selecione uma macro antes de executar."
        GoTo Done
    End If
    If Not TryGetTargetPresentation(target) Then
        Notify "Operacao nao executada: nenhuma apresentacao ativa."
        GoTo Done
    End If
    On Error GoTo Failed
    running = True
    DispatchCommand CStr(cbo.List(cbo.ListIndex, 1))
Done:
    running = False
    MacroLauncherRunSelected = ConsumeNotification()
    If Len(MacroLauncherRunSelected) = 0 Then
        MacroLauncherRunSelected = "Operacao finalizada sem mensagem de resultado."
    End If
    Exit Function
Failed:
    errorMessage = "Falha ao executar a macro: " & Err.Description
    Notify errorMessage
    Resume Done
End Function
