' ------------------------------------------------------------------------------
' (c) balarabe@protonmail.com                                       License: MIT
' :v: 2018-06-02 19:45:30 4962F6                                [sql_insert.bas]
' ------------------------------------------------------------------------------

Option Explicit: Option Compare Text

' sqlInsert reads a row of data from the specified one-row range,
' and returns an SQL INSERT statement as a string.
'
' This function is intended to make extracting data from a worksheet easier.
' It has been used with SQLite & Postgres, but should also work with other DBs.
'
' Note: dates are formatted as 'YYYY-MM-DD' values.
'
' - Use this function to create a formula in an available column
'   (to the left or right of the data you want to extract).
'
' - Copy the formula to all rows you want extract
'   (within the same column used in previous step).
'
' - Select the entire column and copy it.
'   This will copy all the SQL INSERT statements.
'
' - Paste the SQL text into your .sql file or directly in your
'   database's SQL query tool, then run the SQL as required.
'
' EXAMPLES:     sqlInsert("target_table", "qty, rate, amount", A2)
'               sqlInsert("target_table", A$1:A$3, A2)
'
' PARAMETERS:
' tableName:    the name of the table into which to insert rows.
'               It is not escaped in any way, so if escaping is required,
'               specify the table's name as is should be inserted in SQL.
'
' columnList:   -   either a string containing a list of column names
'                   with column names separated by commas
'                   (e.g. "qty, rate, amount")
'
'               -   or a range that contains the column names
'                   (e.g. A$1:C$1)
'
'               To exclude a column, leave the column name blank
'               or wrap the column name in parentheses: (name)
'
'               As with the table's name, column names are not
'               escaped or quoted, so you may need to specify
'               the names as they should be inserted.
'
' firstCell:    The address of the cell from with to start reading data.
'               Additional columns will be read from the cells to
'               the right of the starting cell the same row.
'               (e.g. A2)

Public Function sqlInsert( _
    ByVal tableName As String, _
    ByVal columnList As Variant, _
    ByVal firstCell As Range) As String

    Dim cols() As String
    Dim i As Integer
    If VBA.VarType(columnList) = vbString Then
        cols = VBA.Split(columnList, ",")

    ElseIf VBA.TypeName(columnList) = "Range" Then
        If columnList.Rows.Count > 1 Then
            sqlInsert = "#ERR columnList arg: too many rows!"
            Exit Function
        End If
        ReDim cols(1 To columnList.Columns.Count)
        For i = 1 To columnList.Columns.Count
            cols(i) = columnList.Cells(1, i).Value
        Next
    Else
        sqlInsert = "#ERR columnList arg: invalid type!)"
        Exit Function
    End If

    Dim minCol    As Integer:  minCol = LBound(cols)
    Dim maxCol    As Integer:  maxCol = UBound(cols)
    Dim colCount  As Integer:  colCount = maxCol - minCol + 1

    Dim sql As String
    sql = "INSERT INTO " & tableName & " ("
    Dim f As Boolean
    For i = minCol To maxCol

        Dim field As String
        field = VBA.Trim(cols(i))
        If field <> vbNullString _
        And VBA.Left$(field, 1) <> "(" _
        And VBA.Right$(field, 1) <> ")" Then

            If f Then
                sql = sql & ", "
            End If
            f = True
            sql = sql & field
        End If
    Next
    sql = sql & ") VALUES ("

    Dim hasV As Boolean
    f = False
    For i = minCol To maxCol
        field = VBA.Trim(cols(i))
        If field <> vbNullString _
        And VBA.Left$(field, 1) <> "(" _
        And VBA.Right$(field, 1) <> ")" Then

            If f Then
                sql = sql & ", "
            End If
            f = True

            Dim cellV As Variant
            cellV = firstCell.Offset(0, i - minCol).Value

            If VBA.VarType(cellV) = vbString And VBA.IsDate(cellV) Then
                cellV = VBA.CDate(cellV)
            End If

            Select Case VBA.VarType(cellV)
                Case vbEmpty, vbNull
                    cellV = "NULL"

                Case vbByte, vbCurrency, vbDecimal, vbDouble, _
                    vbInteger, vbLong, vbSingle, vbDouble
                    hasV = True
                    cellV = "" & cellV
                    If VBA.InStrB(1, cellV, ".") > 0 Then
                        While VBA.Right$(cellV, 1) = "0"
                            cellV = VBA.Left$(cellV, VBA.Len(cellV) - 1)
                        Wend
                        While VBA.Right$(cellV, 1) = "."
                            cellV = VBA.Left$(cellV, VBA.Len(cellV) - 1)
                        Wend
                    End If

                Case vbDate
                    hasV = True
                    Dim n As Double
                    n = VBA.CDate(cellV)
                    Dim dateV As Double
                    dateV = VBA.Int(n)
                    cellV = "'"
                    If dateV <> 0 Then
                        cellV = cellV & VBA.Format$(n, "YYYY-MM-DD")
                    End If
                    If (n - dateV) <> 0 Then
                        If dateV <> 0 Then
                            cellV = cellV & " "
                        End If
                        cellV = cellV & VBA.Format$(n, "HH:NN:SS")
                    End If
                    cellV = cellV & "'"

                Case vbString
                    hasV = True
                    cellV = "'" & VBA.Replace$(cellV, "'", "''") & "'"

                Case Else
                    sql = sql & "'#ERR " & VBA.VarType(cellV) & "'"
                    cellV = "NULL"
            End Select
            sql = sql & cellV
        End If
    Next
    sql = sql & ");"
    If Not hasV Then
        sql = vbNullString
    End If
    Debug.Print sql
    sqlInsert = sql
End Function

' eof
