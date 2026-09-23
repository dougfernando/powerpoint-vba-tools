Attribute VB_Name = "modCommandScope"
' @ManagedByDouglasTools
Option Explicit

Public Const COMMAND_SCOPE_SELECTION As String = "selection"
Public Const COMMAND_SCOPE_SLIDE As String = "slide"
Public Const COMMAND_SCOPE_PRESENTATION As String = "presentation"

Private activeCommandScope As String

Public Sub SetCommandScope(ByVal scopeName As String)
    Select Case LCase$(Trim$(scopeName))
        Case COMMAND_SCOPE_SELECTION, COMMAND_SCOPE_SLIDE, COMMAND_SCOPE_PRESENTATION
            activeCommandScope = LCase$(Trim$(scopeName))
        Case Else
            Err.Raise 5, "DFS Tools", "Escopo de comando invalido: " & scopeName
    End Select
End Sub

Public Sub ResetCommandScope()
    activeCommandScope = COMMAND_SCOPE_SLIDE
End Sub

Public Function CurrentCommandScope() As String
    If Len(activeCommandScope) = 0 Then ResetCommandScope
    CurrentCommandScope = activeCommandScope
End Function

Public Function CommandScopeLabel(ByVal scopeName As String) As String
    Select Case LCase$(Trim$(scopeName))
        Case COMMAND_SCOPE_SELECTION
            CommandScopeLabel = "selecao"
        Case COMMAND_SCOPE_PRESENTATION
            CommandScopeLabel = "apresentacao"
        Case Else
            CommandScopeLabel = "slide"
    End Select
End Function
