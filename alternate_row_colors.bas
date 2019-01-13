Option Explicit: Option Compare Text

Sub alternateRowColors()
    Dim cel As Range
    For Each cel In Application.Selection
        Dim isEven As Boolean
        isEven = cel.Row Mod 2 = 0
        With cel.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .PatternTintAndShade = 0
            .ThemeColor = VBA.IIf(isEven, xlThemeColorAccent1, xlThemeColorDark1)
            .TintAndShade = VBA.IIf(isEven, 0.799981688894314, -4.99893185216834E-02)
        End With
    Next
End Sub

' eof
