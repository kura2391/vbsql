Public Class Delete
    Private _from As String
    Private _where As String

    'set from
    Public Function from(table As String) As Delete
        _from = table
        Return Me
    End Function

    ' set where
    Public Function where(conditions As String) As Delete
        _where = conditions
        Return Me
    End Function

    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        checkFrom()

        Dim cn As New SqlClient.SqlConnection
        Dim sql As New SqlClient.SqlCommand
        cn.ConnectionString = connectionString
        sql = cn.CreateCommand
        sql.CommandText = getSql()
        cn.Open()

        execute = sql.ExecuteScalar()

        cn.Close()
        sql.Dispose()
        cn.Dispose()
    End Function

    Function getSql()
        Dim sql As String = " DELETE FROM "
        sql &= _from
        If Not IsNothing(_where) Then
            sql &= " WHERE "
            sql &= _where
        End If
        Return sql
    End Function


    'check from
    Public Sub checkFrom()
        If IsNothing(_from) OrElse _from.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub


End Class
