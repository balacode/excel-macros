' ------------------------------------------------------------------------------
' (c) balarabe@protonmail.com                                       License: MIT
' :v: 2018-05-31 15:53:00 DD00E8                                 [sql_insert.go]
' ------------------------------------------------------------------------------

Option Explicit: Option Compare Text

Public Function sqlInsert( _
    ByVal tableName As String, _
    ByVal columnList As String, _
    ByVal firstCol As Range) As String

    Dim cols      As Variant:  cols = VBA.Split(columnList, ",")
    Dim minCol    As Integer:  minCol = LBound(cols)
    Dim maxCol    As Integer:  maxCol = UBound(cols)
    Dim colCount  As Integer:  colCount = maxCol - minCol + 1

    Dim sql As String
    sql = "INSERT INTO " & tableName & " ("
    Dim i As Integer
    For i = minCol To maxCol
        If i <> minCol Then
            sql = sql & ", "
        End If
        sql = sql & Trim(cols(i))
    Next
    sql = sql & ") VALUES ("

    Dim hasValue As Boolean

    For i = minCol To maxCol
        If i <> minCol Then
            sql = sql & ", "
        End If

        Dim cellVal As Variant
        cellVal = firstCol.Offset(0, i - minCol).Value

        If VBA.VarType(cellVal) = vbString And VBA.IsDate(cellVal) Then
            cellVal = VBA.CDate(cellVal)
        End If

        Select Case VBA.VarType(cellVal)
            Case vbEmpty, vbNull
                cellVal = "NULL"

            Case vbByte, vbCurrency, vbDecimal, vbDouble, _
                vbInteger, vbLong, vbSingle, vbDouble
                hasValue = True
                cellVal = "" & cellVal
                If VBA.InStrB(1, cellVal, ".") > 0 Then
                    While VBA.Right$(cellVal, 1) = "0"
                        cellVal = VBA.Left$(cellVal, VBA.Len(cellVal) - 1)
                    Wend
                    While VBA.Right$(cellVal, 1) = "."
                        cellVal = VBA.Left$(cellVal, VBA.Len(cellVal) - 1)
                    Wend
                End If
            Case vbDate
                hasValue = True
                Dim dt As Double
                Dim dateVal As Double
                dt = VBA.CDate(cellVal)
                dateVal = VBA.Int(dt)
                cellVal = "'"
                If dateVal <> 0 Then
                    cellVal = cellVal & VBA.Format$(dt, "YYYY-MM-DD")
                End If
                If (dt - dateVal) <> 0 Then
                    If dateVal <> 0 Then
                        cellVal = cellVal & " "
                    End If
                    cellVal = cellVal & VBA.Format$(dt, "HH:NN:SS")
                End If
                cellVal = cellVal & "'"
            Case vbString
                hasValue = True
                cellVal = "'" & VBA.Replace$(cellVal, "'", "''") & "'"
            Case vbArray:           cellVal = "'vbArray'"
            Case vbBoolean:         cellVal = "'vbBoolean'"
            Case vbDataObject:      cellVal = "'vbDataObject'"
            Case vbError:           cellVal = "'vbError'"
            Case vbObject:          cellVal = "'vbObject'"
            Case vbUserDefinedType: cellVal = "'vbUserDefinedType'"
            Case vbVariant:         cellVal = "'vbVariant'"
            Case Else
                sql = sql & "'TYPE" & VBA.VarType(cellVal) & "'"
        End Select
        sql = sql & cellVal
    Next
    sql = sql & ");"
    If hasValue Then
        sqlInsert = sql
    Else
        sqlInsert = vbNullString
    End If
End Function

' eof
