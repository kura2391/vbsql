Public Class Delete
    Inherits SqlAbstract


    Private _where As New Where("where")

    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub
    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub


    'set from
    Public Function from(table As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As Delete
        _table.Value = table
        _table.DbType = type
        Return Me
    End Function

    ' set where
    Public Function where(conditions As String, col As String()) As Delete
        _where.set(conditions, col)
        Return Me
    End Function

    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        check()
        Return _connection.execute(buildSql(), buildParameter())
    End Function


    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.Add(_table)
        p.AddRange(_where.getParameterList())
        Return p.ToArray()
    End Function

    Private Function buildSql()
        Dim sql As String = " DELETE FROM @"
        sql &= _table.ParameterName

        If Not _where.isEmpty() Then
            sql &= " WHERE "
            sql &= _where.sql()
        End If
        Return sql
    End Function

    Public Sub check()
        checkConnection()
    End Sub



End Class
