Attribute VB_Name = "modMacroLauncher"
Option Explicit
' @ManagedByDouglasTools

Private launcher As frmMacroLauncher
Private running As Boolean
Private closingLauncher As Boolean

Public Sub ShowMacroLauncher()
    Dim firstError As String

    On Error GoTo RetryWithFreshInstance
    ShowLauncherInstance
    Exit Sub
RetryWithFreshInstance:
    firstError = Err.Description
    On Error GoTo Failed
    CloseMacroLauncher
    ShowLauncherInstance
    Exit Sub
Failed:
    Notify "Nao foi possivel abrir as macros: " & Err.Description & _
           ". Primeira tentativa: " & firstError
End Sub

Private Sub ShowLauncherInstance()
    If Not launcher Is Nothing Then
        If Not launcher.Visible Then CloseMacroLauncher
    End If
    If launcher Is Nothing Then Set launcher = New frmMacroLauncher
    launcher.PrepareForDisplay
    launcher.Show vbModeless
End Sub

Public Sub RibbonShowMacroLauncher(ByVal control As Office.IRibbonControl)
    ShowMacroLauncher
End Sub

Public Sub CloseMacroLauncher()
    Dim target As frmMacroLauncher
    Dim errorNumber As Long
    Dim errorDescription As String

    If closingLauncher Then Exit Sub
    closingLauncher = True

    On Error GoTo Failed
    CancelNotificationClear
    CancelPendingConfirmation
    Set target = launcher
    Set launcher = Nothing
    If Not target Is Nothing Then Unload target
Done:
    closingLauncher = False
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    If launcher Is Nothing Then Set launcher = target
    closingLauncher = False
    Err.Raise errorNumber, "DFS Tools", errorDescription
End Sub

Public Sub ReleaseMacroLauncher(ByVal terminatedLauncher As frmMacroLauncher)
    If launcher Is Nothing Then Exit Sub
    If launcher Is terminatedLauncher Then Set launcher = Nothing
End Sub

Public Sub ClearLauncherResult()
    On Error Resume Next
    If Not launcher Is Nothing Then launcher.ClearResult
    On Error GoTo 0
End Sub

Public Sub ShowLauncherResult(ByVal message As String, ByVal state As String)
    On Error Resume Next
    ShowMacroLauncher
    If Not launcher Is Nothing Then
        launcher.DisplayResult message, state
        ScheduleNotificationClear
    End If
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
    cbo.ColumnCount = 4
    cbo.ColumnWidths = CStr(cbo.Width - 22) & " pt;0 pt;0 pt;0 pt"
    For Each item In CommandCatalog()
        If category = "Todas" Or category = CStr(item(1)) Then
            cbo.AddItem CStr(item(2)) & " (" & CStr(item(1)) & ")"
            cbo.List(cbo.ListCount - 1, 1) = CStr(item(0))
            cbo.List(cbo.ListCount - 1, 2) = CStr(item(3))
            cbo.List(cbo.ListCount - 1, 3) = CStr(item(4))
        End If
    Next item
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
End Sub

Public Function MacroLauncherRunSelected(ByVal cbo As Object, ByVal scopeName As String) As String
    Dim target As Presentation
    Dim errorMessage As String
    Dim scopeSpec As String

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
    scopeSpec = CStr(cbo.List(cbo.ListIndex, 3))
    If Len(scopeSpec) > 0 Then
        If Not CommandSupportsScope(scopeSpec, scopeName) Then
            Notify "Operacao nao executada: o escopo escolhido nao e valido para esta macro."
            GoTo Done
        End If
        SetCommandScope scopeName
    Else
        ResetCommandScope
    End If
    On Error GoTo Failed
    running = True
    DispatchCommand CStr(cbo.List(cbo.ListIndex, 1))
Done:
    running = False
    MacroLauncherRunSelected = ConsumeNotification()
    ResetCommandScope
    If Len(MacroLauncherRunSelected) = 0 Then
        MacroLauncherRunSelected = "Operacao finalizada sem mensagem de resultado."
    End If
    Exit Function
Failed:
    errorMessage = "Falha ao executar a macro: " & Err.Description
    Notify errorMessage
    Resume Done
End Function

Public Function CommandSupportsScope(ByVal scopeSpec As String, ByVal scopeName As String) As Boolean
    Dim normalizedSpec As String
    normalizedSpec = "," & LCase$(Replace(scopeSpec, " ", vbNullString)) & ","
    CommandSupportsScope = (InStr(1, normalizedSpec, "," & LCase$(scopeName) & ",", vbBinaryCompare) > 0)
End Function
