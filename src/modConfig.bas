Attribute VB_Name = "modConfig"
' @ManagedByDouglasTools
Option Explicit

Public Const CM_TO_POINTS As Double = 28.3464567
Public Const STANDARD_MARGIN_CM As Double = 0.2
Public Const SHAPE_MATCH_TOLERANCE_PT As Double = 0.1
Public Const QA_MARKER_TAG_NAME As String = "DOUGLAS_QA_MARKER"
Public Const QA_MARKER_FONT As String = "FONT"
Public Const QA_MARKER_CLIENT As String = "CLIENT"

Public Function StandardMarginPoints() As Single
    StandardMarginPoints = CSng(STANDARD_MARGIN_CM * CM_TO_POINTS)
End Function

Public Function AllowedFontPrefixes() As Variant
    AllowedFontPrefixes = Array("Graphik", "GT Sectra")
End Function

Public Function ForbiddenClientTerms() As Variant
    ' Edit this list as needed. Short terms should generally be checked as whole words only.
    ForbiddenClientTerms = Array( _
        "Santander", "PagoNxt", "Getnet", "Zurich", "Pefisa", "Cielo", _
        "Bradesco", "ABECS", "Elo", "Cateno", "Itaú", "XP", "Itausa", _
        "Porto Seguro", "Redecard", "BV", "Brasilseg", "BB", "Mapfre", _
        "C6", "Sulamerica", "Banco ABC", "Prudential", "Coruja", "BTG", _
        "Chicago", "Google", "Boa Vista", "Febraban", "Safra", "ANBIMA", _
        "B3", "Nubank", "NU", "Chubb", "Sodexo", "Market Pay", "CIP", _
        "BNY", "Ticket", "Allianz", "PayPal", "Banco Sofisa", "CERC", _
        "Stone", "Icatu", "Serasa", "Banco Honda", "Sicoob", "SPX" _
    )
End Function
