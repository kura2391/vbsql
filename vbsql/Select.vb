
Public Class [Select]
        Private _columns As String = "*"
        Private _from As String = Nothing
        Private _where As String = Nothing
        Private _order As String = Nothing

        'set table
        Public Function from(table As String) As [Select]
            _from = table
            Return Me
        End Function
        'set Columns
        Public Function [select](columns As String) As [Select]
            _columns = columns
            Return Me
        End Function
        'set Where
        Public Function where(conditions As String) As [Select]
            _where = conditions
            Return Me
        End Function

        'set OrderBy
        Public Function orderBy(order As String) As [Select]
            _order = order
            Return Me
        End Function

        'execute sql and get data as DataTable
        Public Function execute(connectionString As String) As DataTable
            checkFrom()
            Dim dt As New DataTable()
            Dim da = New SqlClient.SqlDataAdapter(Me.getSql, connectionString)
            da.Fill(dt)
            Return dt
        End Function

        'create sql sentence 
        Private Function getSql() As String
            Dim sql As String = "SELECT "
            sql &= _columns
            sql &= " FROM "
            sql &= _from
            If Not IsNothing(_where) Then
                sql &= " WHERE "
                sql &= _where
            End If
            If Not IsNothing(_order) Then
                sql &= " ORDER BY "
                sql &= _order
            End If
            Return sql
        End Function

        'if _from is not set, throw error
        Private Sub checkFrom()
            If IsNothing(_from) OrElse _from.Trim() = "" Then
                Throw New Exception("no table selected")
            End If
        End Sub
End Class