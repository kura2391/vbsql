Public Class Delete
    Inherits SqlAbstract


    Private _where As New Where()

    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub
    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub


    'set from
    Public Function from(table As String) As Delete
        _table = table
        Return Me
    End Function

    ' set where
    Public Function where(conditions As String, col As Parameter()) As Delete
        _where.add(conditions, col)
        Return Me
    End Function
    ' set where
    Public Function where(conditions As String, col() As String) As Delete
        _where.add(conditions, col)
        Return Me
    End Function

    'execute sql and return integer 
    Public Function execute() As Integer
        check()
        Return _connection.execute(buildSql(), buildParameter())
    End Function


    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.AddRange(_where.getParamList())
        Return p.ToArray()
    End Function

    Private Function buildSql()
        Dim sql As String = " DELETE FROM "
        sql &= _table

        If Not _where.isEmpty() Then
            sql &= " WHERE "
            sql &= _where.sql()
        End If
        Return sql
    End Function


    Private Sub check()
        checkConnection()
    End Sub



End Class
