Public Class Update
    Inherits SqlAbstract
    Private _set As New Parameter("set")
    Private _where As New Where("where")

    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub
    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub


    Public Function update(table As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As Update
        _table.Value = table
        _table.DbType = type
        Return Me
    End Function

    Public Function [set](columnName As String, value As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As Update
        _set.appendParameter(columnName)
        _set.appendParameter(value, type)
        Return Me
    End Function

    'set where
    Public Function where(conditions As String, col() As String) As Update
        _where.set(conditions, col)
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        check()
        Return _connection.execute(buildSql(), buildParameter())
    End Function

    '
    Function buildSql() As String
        Dim sql As String = "UPDATE @" & _table.ParameterName & " SET "
        For i As Integer = 0 To _set.count() - 1
            sql &= "@" & _set.getParameter(i).ParameterName & "=@" & _set.getParameter(i).Value & ","
        Next
        sql = sql.Substring(0, sql.Length - 1)
        If Not _where.isEmpty() Then
            sql &= " WHERE " & _where.sql()
        End If
        Return sql
    End Function

    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.add(_table)
        p.AddRange(_set.getParameterList)
        p.AddRange(_where.getParameterList)

        Return p.ToArray()
    End Function

    Public Sub check()
        checkUpdate()
        checkSet()
        checkConnection()
    End Sub

    'if _from is not set, throw error
    Sub checkUpdate()
        If IsNothing(_table.Value) Then
            Throw New Exception("no table selected")
        End If
    End Sub

    'check _set
    Sub checkSet()
        If IsNothing(_set) OrElse _set.count = 0 Then
            Throw New Exception("no data set")
        End If
    End Sub

End Class
