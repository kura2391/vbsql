Public Class Update
    Private _update As String = Nothing
    Private _set As Hashtable = Nothing
    Private _where As String = Nothing

    Public Function update(table As String) As Update
        _update = table
        Return Me
    End Function

    Public Function [set](hashtable As Hashtable) As Update
        _set = hashtable
        Return Me
    End Function

    'set where
    Public Function where(conditions As String) As Update
        _where = conditions
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        checkUpdate()
        checkSet()

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

    '
    Function getSql() As String
        Dim sql As String = ""
        sql = "UPDATE "
        sql &= _update
        sql &= " SET "
        For Each key As String In _set.Keys
            sql &= key & "='"
            sql &= _set(key) & "',"
        Next
        sql = sql.Substring(0, sql.Length - 1)
        If Not IsNothing(_where) Then
            sql &= " WHERE "
            sql &= _where
        End If

        Return sql
    End Function





    'if _from is not set, throw error
    Sub checkUpdate()
        If IsNothing(_update) OrElse _update.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub

    'check _set
    Sub checkSet()
        If IsNothing(_set) OrElse _set.Count = 0 Then
            Throw New Exception("no data set")
        End If
    End Sub

End Class
