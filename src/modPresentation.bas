Attribute VB_Name = "modPresentation"
' @ManagedByDouglasTools
Option Explicit

Public Function TryGetTargetPresentation(ByRef pres As Presentation) As Boolean
    On Error GoTo Fail

    If Application.Presentations.Count = 0 Then
        Exit Function
    End If

    Set pres = ActivePresentation
    TryGetTargetPresentation = Not pres Is Nothing
    Exit Function

Fail:
    TryGetTargetPresentation = False
End Function

Public Function TryGetActiveSlide(ByRef sld As Slide) As Boolean
    On Error GoTo Fail

    Set sld = ActiveWindow.View.Slide
    TryGetActiveSlide = Not sld Is Nothing
    Exit Function

Fail:
    TryGetActiveSlide = False
End Function

Public Function SlideWidthPoints(ByVal pres As Presentation) As Single
    SlideWidthPoints = pres.PageSetup.SlideWidth
End Function

Public Function SlideHeightPoints(ByVal pres As Presentation) As Single
    SlideHeightPoints = pres.PageSetup.SlideHeight
End Function
