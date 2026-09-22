Attribute VB_Name = "modNotifications"
' @ManagedByDouglasTools
Option Explicit

Private lastNotification As String
Private pendingConfirmationKey As String
Private pendingConfirmationDeadline As Date

#If VBA7 Then
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal eventId As LongPtr, ByVal intervalMs As Long, ByVal callbackAddress As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal eventId As LongPtr) As Long
Private clearTimerId As LongPtr
#Else
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal eventId As Long, ByVal intervalMs As Long, ByVal callbackAddress As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal eventId As Long) As Long
Private clearTimerId As Long
#End If

Private Const RESULT_DISPLAY_TIME_MS As Long = 1000

Public Sub Notify(ByVal message As String)
    lastNotification = Trim$(message)
End Sub

Public Sub ClearNotification()
    lastNotification = vbNullString
End Sub

Public Function ConsumeNotification() As String
    ConsumeNotification = lastNotification
    lastNotification = vbNullString
End Function

Public Sub ScheduleNotificationClear()
    CancelNotificationClear
    clearTimerId = SetTimer(0, 0, RESULT_DISPLAY_TIME_MS, AddressOf NotificationTimerCallback)
End Sub

Public Sub CancelNotificationClear()
    If clearTimerId <> 0 Then
        KillTimer 0, clearTimerId
        clearTimerId = 0
    End If
End Sub

Public Function ConfirmThroughStatus(ByVal confirmationKey As String, ByVal prompt As String) As Boolean
    If StrComp(pendingConfirmationKey, confirmationKey, vbBinaryCompare) = 0 And _
       pendingConfirmationDeadline >= Now Then
        CancelPendingConfirmation
        ConfirmThroughStatus = True
        Exit Function
    End If

    pendingConfirmationKey = confirmationKey
    pendingConfirmationDeadline = DateAdd("s", 1, Now)
    Notify prompt & " Execute novamente em ate 1 segundo para confirmar."
End Function

Public Sub CancelPendingConfirmation()
    pendingConfirmationKey = vbNullString
    pendingConfirmationDeadline = 0
End Sub

#If VBA7 Then
Public Sub NotificationTimerCallback(ByVal hWnd As LongPtr, ByVal message As Long, ByVal eventId As LongPtr, ByVal tickCount As Long)
#Else
Public Sub NotificationTimerCallback(ByVal hWnd As Long, ByVal message As Long, ByVal eventId As Long, ByVal tickCount As Long)
#End If
    CancelNotificationClear
    CancelPendingConfirmation
    modMacroLauncher.ClearLauncherResult
End Sub
