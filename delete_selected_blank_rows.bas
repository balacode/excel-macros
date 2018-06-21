' ------------------------------------------------------------------------------
' (c) balarabe@protonmail.com                                       License: MIT
' :v: 2018-06-21 20:48:48 75A6ED   excel-macros/[delete_selected_blank_rows.bas]
' ------------------------------------------------------------------------------

Option Explicit: Option Compare Text

' deletes all selected completely-blank rows in the current worksheet
Public Sub deleteSelectedBlankRows()

    Dim ws As Worksheet:    Set ws = Application.ActiveSheet

    Dim rowCount As Long:   rowCount = Selection.Rows.Count
    Dim colCount As Long:   colCount = Selection.Columns.Count

    Dim firstRow As Long:   firstRow = Selection.Rows(1).row
    Dim lastRow As Long:    lastRow = Selection.Rows(rowCount).row

    Dim firstCol As Long:   firstCol = Selection.Columns(1).Column
    Dim lastCol As Long:    lastCol = Selection.Columns(colCount).Column

    Dim row   As Long
    For row = lastRow To firstRow Step -1

        Dim isBlank As Boolean
        isBlank = True

        Dim col As Long
        For col = firstCol To lastCol
            If Not VBA.IsEmpty(ws.Cells(row, col).Value) Then
                isBlank = False
                Exit For
            End If
        Next

        If isBlank Then
            ws.Cells(row, 1).EntireRow.Delete
        End If
    Next
End Sub '                                                deleteSelectedBlankRows

' end
