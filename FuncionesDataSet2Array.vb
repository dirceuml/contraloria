Module FuncionesDataSet2Array

    Function DataSet2Array(ByVal dt As DataTable, ByVal flagTitulos As Boolean) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(dt.Rows.Count, dt.Columns.Count - 1) As Object

        If flagTitulos Then
            'Copy the column names to the first row of the object array
            For col = 0 To dt.Columns.Count - 1
                rawData(0, col) = dt.Columns(col).ColumnName
            Next
        End If
        ' Copy the values to the object array
        For col = 0 To dt.Columns.Count - 1
            For row = 0 To dt.Rows.Count - 1
                If Not IsDBNull(dt.Rows(row).ItemArray(col)) Then
                    If flagTitulos Then rawData(row + 1, col) = dt.Rows(row).ItemArray(col) Else rawData(row, col) = dt.Rows(row).ItemArray(col)
                End If
            Next
        Next

        Return rawData
    End Function

    Function DataSet2Array(ByVal dt As DataTable, ByVal row_ini As Integer, ByVal row_fin As Integer) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(row_fin - row_ini, dt.Columns.Count - 1) As Object

        'Copy the column names to the first row of the object array
        'For col = 0 To dt.Columns.Count - 1
        '    rawData(0, col) = dt.Columns(col).ColumnName
        'Next

        ' Copy the values to the object array
        For col = 0 To dt.Columns.Count - 1
            For i = 0 To row_fin - row_ini
                If Not IsDBNull(dt.Rows(row_ini + i).ItemArray(col)) Then
                    rawData(i, col) = dt.Rows(row_ini + i).ItemArray(col)
                End If
            Next
        Next

        Return rawData
    End Function

    Function DataSet2Array(ByVal dt As DataTable, ByVal CantCol As Integer, ByVal col1 As Integer, ByVal col2 As Integer, ByVal col3 As Integer, ByVal col4 As Integer, ByVal col5 As Integer) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(dt.Rows.Count - 1, CantCol - 1) As Object

        'Copy the column names to the first row of the object array
        'For col = 0 To dt.Columns.Count - 1
        '    rawData(0, col) = dt.Columns(col).ColumnName
        'Next

        ' Copy the values to the object array
        For row = 0 To dt.Rows.Count - 1
            If Not IsDBNull(dt.Rows(row).ItemArray(col1)) Then rawData(row, 0) = dt.Rows(row).ItemArray(col1)
            If col2 > -1 Then
                If Not IsDBNull(dt.Rows(row).ItemArray(col2)) Then rawData(row, 1) = dt.Rows(row).ItemArray(col2)
            End If
            If col3 > -1 Then
                If Not IsDBNull(dt.Rows(row).ItemArray(col3)) Then rawData(row, 2) = dt.Rows(row).ItemArray(col3)
            End If
            If col4 > -1 Then
                If Not IsDBNull(dt.Rows(row).ItemArray(col4)) Then rawData(row, 3) = dt.Rows(row).ItemArray(col4)
            End If
            If col5 > -1 Then
                If Not IsDBNull(dt.Rows(row).ItemArray(col5)) Then rawData(row, 4) = dt.Rows(row).ItemArray(col5)
            End If
        Next

        Return rawData
    End Function

    Function DataSet2Array(ByVal dt As DataTable, ByVal row_ini As Integer, ByVal row_fin As Integer, ByVal CantCol As Integer, ByVal col1 As Integer, ByVal col2 As Integer, ByVal col3 As Integer, ByVal col4 As Integer, ByVal col5 As Integer) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(row_fin - row_ini, CantCol - 1) As Object
        'Copy the column names to the first row of the object array
        'For col = 0 To dt.Columns.Count - 1
        '    rawData(0, col) = dt.Columns(col).ColumnName
        'Next

        ' Copy the values to the object array
        For row = 0 To row_fin - row_ini
            If Not IsDBNull(dt.Rows(row_ini + row).ItemArray(col1)) Then rawData(row, 0) = dt.Rows(row_ini + row).ItemArray(col1)
            If col2 > -1 Then
                If Not IsDBNull(dt.Rows(row_ini + row).ItemArray(col2)) Then rawData(row, 1) = dt.Rows(row_ini + row).ItemArray(col2)
            End If
            If col3 > -1 Then
                If Not IsDBNull(dt.Rows(row_ini + row).ItemArray(col3)) Then rawData(row, 2) = dt.Rows(row_ini + row).ItemArray(col3)
            End If
            If col4 > -1 Then
                If Not IsDBNull(dt.Rows(row_ini + row).ItemArray(col4)) Then rawData(row, 3) = dt.Rows(row_ini + row).ItemArray(col4)
            End If
            If col5 > -1 Then
                If Not IsDBNull(dt.Rows(row_ini + row).ItemArray(col5)) Then rawData(row, 4) = dt.Rows(row_ini + row).ItemArray(col5)
            End If
        Next

        Return rawData
    End Function

    Function DataSet2Array(ByVal dr As DataRow(), ByVal CantCol As Integer, ByVal col1 As Integer, ByVal col2 As Integer, ByVal col3 As Integer, ByVal col4 As Integer, ByVal col5 As Integer) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(dr.GetUpperBound(0) + 1, dr(0).ItemArray.GetUpperBound(0) + 1) As Object
        'Copy the column names to the first row of the object array
        'For col = 0 To dt.Columns.Count - 1
        '    rawData(0, col) = dt.Columns(col).ColumnName
        'Next

        ' Copy the values to the object array
        For row = 0 To dr.GetUpperBound(0)
            If Not IsDBNull(dr(row).ItemArray(col1)) Then rawData(row, 0) = dr(row).ItemArray(col1)
            If col2 > -1 Then
                If Not IsDBNull(dr(row).ItemArray(col2)) Then rawData(row, 1) = dr(row).ItemArray(col2)
            End If
            If col3 > -1 Then
                If Not IsDBNull(dr(row).ItemArray(col3)) Then rawData(row, 2) = dr(row).ItemArray(col3)
            End If
            If col4 > -1 Then
                If Not IsDBNull(dr(row).ItemArray(col4)) Then rawData(row, 3) = dr(row).ItemArray(col4)
            End If
            If col5 > -1 Then
                If Not IsDBNull(dr(row).ItemArray(col5)) Then rawData(row, 4) = dr(row).ItemArray(col5)
            End If
        Next

        Return rawData
    End Function

    Function DataSet2Array(ByVal dr As DataRow(), ByVal row_ini As Integer, ByVal row_fin As Integer, ByVal CantCol As Integer, ByVal col1 As Integer, ByVal col2 As Integer, ByVal col3 As Integer, ByVal col4 As Integer, ByVal col5 As Integer) As Object(,)
        ' Copy the DataTable to an object array
        Dim rawData(row_fin - row_ini, dr(0).ItemArray.GetUpperBound(0) + 1) As Object
        'Copy the column names to the first row of the object array
        'For col = 0 To dt.Columns.Count - 1
        '    rawData(0, col) = dt.Columns(col).ColumnName
        'Next

        ' Copy the values to the object array
        For row = 0 To row_fin - row_ini
            If Not IsDBNull(dr(row_ini + row).ItemArray(col1)) Then rawData(row, 0) = dr(row_ini + row).ItemArray(col1)
            If col2 > -1 Then
                If Not IsDBNull(dr(row_ini + row).ItemArray(col2)) Then rawData(row, 1) = dr(row_ini + row).ItemArray(col2)
            End If
            If col3 > -1 Then
                If Not IsDBNull(dr(row_ini + row).ItemArray(col3)) Then rawData(row, 2) = dr(row_ini + row).ItemArray(col3)
            End If
            If col4 > -1 Then
                If Not IsDBNull(dr(row_ini + row).ItemArray(col4)) Then rawData(row, 3) = dr(row_ini + row).ItemArray(col4)
            End If
            If col5 > -1 Then
                If Not IsDBNull(dr(row_ini + row).ItemArray(col5)) Then rawData(row, 4) = dr(row_ini + row).ItemArray(col5)
            End If
        Next

        Return rawData
    End Function

End Module
