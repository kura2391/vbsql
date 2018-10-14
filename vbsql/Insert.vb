Public Class Insert
    Inherits SqlAbstract

    'save lastinsertid
    Private _lastInsertId As Boolean = False
    Private _value As New Parameter("value")


    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub

    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub


    'set table
    Public Function into(table As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As Insert
        _table.Value = table
        _table.DbType = type
        Return Me
    End Function

    'set insertdata
    Public Function values(columnName As String, value As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As Insert
        _value.appendParameter(columnName)
        _value.appendParameter(value, type)
        Return Me
    End Function

    'set lastInsertId
    Public Function lastInsertId() As Insert
        _lastInsertId = True
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute(connectionString As String) As Integer
        check()

        Dim sql As String = buildSql()
        Return _connection.execute(sql, buildParameter())
    End Function

    'create sql
    Private Function buildSql()
        'append table parameter to the list of sqlParameter 

        Dim sql As String = "INSERT INTO @" & _table.ParameterName & " ("
        Dim values As String = ""
        For i As Integer = 0 To _value.getParameterList.Count - 2 Step 2
            sql &= "@" & _value.getParameter(i).ParameterName & ","
            values &= "@" & _value.getParameter(i + 1).ParameterName & ","
        Next
        sql = sql.Substring(0, sql.Length - 1) & ") VALUES("
        sql &= values.Substring(0, values.Length - 1) & ")"

        If _lastInsertId Then
            sql &= "SELECT SCOPE_IDENTITY();"
        End If
        Return sql
    End Function

    Public Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.Add(_table)
        p.AddRange(_value.getParameterList)
        Return p.ToArray()
    End Function

    Private Sub check()
        checkTable()
        checkValues()
        MyBase.checkConnection()
    End Sub


    'check table
    Private Sub checkTable()
        If IsNothing(_table) Then
            Throw New Exception("no table selected")
        End If
    End Sub

    Private Sub checkValues()
        If _value.getParameterList.Count < 2 Then
            Throw New Exception("no data inserted")
        End If
    End Sub

End Class
