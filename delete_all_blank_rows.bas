' ------------------------------------------------------------------------------
' (c) balarabe@protonmail.com                                       License: MIT
' :v: 2018-06-20 09:25:49 079A98        excel-macros/[delete_all_blank_rows.bas]
' ------------------------------------------------------------------------------

Option Explicit: Option Compare Text

' deletes all completely-blank rows in the current worksheet
Public Sub deleteAllBlankRows()

    Dim ws As Worksheet:    Set ws = Application.ActiveSheet
    Dim lastRow As Long:    lastRow = ws.UsedRange.Rows.count
    Dim lastCol As Long:    lastCol = ws.UsedRange.Columns.count

    Dim row As Long
    For row = lastRow To 1 Step -1

        Dim isBlank As Boolean
        isBlank = True

        Dim col As Long
        For col = 1 To lastCol

            If Not VBA.IsEmpty(ws.Cells(row, col).Value) Then
                isBlank = False
                Exit For
            End If
        Next

        If isBlank Then
            ws.Cells(row, 1).EntireRow.Delete
        End If
    Next
End Sub '                                                     deleteAllBlankRows

' end
