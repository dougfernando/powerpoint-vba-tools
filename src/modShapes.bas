Attribute VB_Name = "modShapes"
' @ManagedByDouglasTools
Option Explicit

Public Function ShapeHasNoFill(ByVal shp As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasNoFill = (shp.Fill.Visible = msoFalse) Or (shp.Fill.Transparency >= 0.99)
    Exit Function
Fail:
    ShapeHasNoFill = False
End Function

Public Function ShapeHasVisibleFill(ByVal shp As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasVisibleFill = (shp.Fill.Visible = msoTrue And shp.Fill.Transparency < 0.99)
    Exit Function
Fail:
    ShapeHasVisibleFill = False
End Function

Public Function ShapeHasText(ByVal shp As Shape) As Boolean
    On Error GoTo Fail
    ShapeHasText = (shp.HasTextFrame = msoTrue And shp.TextFrame2.HasText = msoTrue)
    Exit Function
Fail:
    ShapeHasText = False
End Function

Public Function ShapeSupportsText(ByVal shp As Shape) As Boolean
    On Error GoTo Fail
    ShapeSupportsText = (shp.HasTextFrame = msoTrue)
    Exit Function
Fail:
    ShapeSupportsText = False
End Function

Public Function SameRect(ByVal shp As Shape, ByVal refL As Double, ByVal refT As Double, ByVal refW As Double, ByVal refH As Double, ByVal tol As Double) As Boolean
    On Error GoTo Fail
    SameRect = (Abs(shp.Left - refL) <= tol) And _
               (Abs(shp.Top - refT) <= tol) And _
               (Abs(shp.Width - refW) <= tol) And _
               (Abs(shp.Height - refH) <= tol)
    Exit Function
Fail:
    SameRect = False
End Function

Public Function ShapeIsCompletelyOutsideSlide(ByVal shp As Shape, ByVal slideWidth As Single, ByVal slideHeight As Single) As Boolean
    On Error GoTo Fail
    ShapeIsCompletelyOutsideSlide = _
        (shp.Left >= slideWidth) Or _
        (shp.Left + shp.Width <= 0) Or _
        (shp.Top >= slideHeight) Or _
        (shp.Top + shp.Height <= 0)
    Exit Function
Fail:
    ShapeIsCompletelyOutsideSlide = False
End Function

Public Function ShapeFillRgb(ByVal shp As Shape, ByRef rgbValue As Long) As Boolean
    On Error GoTo Fail
    If shp.Fill.Visible <> msoTrue Then Exit Function
    rgbValue = shp.Fill.ForeColor.RGB
    ShapeFillRgb = True
    Exit Function
Fail:
    ShapeFillRgb = False
End Function
