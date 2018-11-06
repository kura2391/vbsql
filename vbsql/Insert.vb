
Public Class Insert
    Inherits SqlAbstract

    'save lastinsertid
    Private _lastInsertId As Boolean = False


    Public Sub New(connection As Connection)
        MyBase.New(connection)
        _prefix = "@insert"
    End Sub

    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
        _prefix = "@insert"
    End Sub


    'set table
    Public Function into(table As String) As Insert
        _table = table
        Return Me
    End Function

    'set insertdata
    Public Function values(ByVal paramList As List(Of Parameter)) As Insert
        _variables = New List(Of Parameter)(paramList)
        Return Me
    End Function

    Public Function values(paramList As Dictionary(Of String, String)) As Insert
        Dim param As New List(Of Parameter)
        For Each key As String In paramList.Keys
            param.Add(New Parameter(key, paramList(key)))
        Next
        Return values(param)
    End Function

    'set lastInsertId
    Public Function lastInsertId() As Insert
        _lastInsertId = True
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute() As Integer
        check()
        setParameterName()
        Return _connection.execute(buildSql(), buildParameter())
    End Function

    'create sql
    Private Function buildSql()

        Dim sql As String = "INSERT INTO " & _table & " ("
        Dim values As String = ""
        For i As Integer = 0 To _variables.Count - 1
            sql &= " " & _variables(i).getColumnName & " ,"
            values &= " " & _variables(i).getParameterName & " ,"
        Next
        sql = sql.Substring(0, sql.Length - 1) & ") VALUES ("
        sql &= values.Substring(0, values.Length - 1) & ")"

        If _lastInsertId Then
            sql &= "SELECT SCOPE_IDENTITY();"
        End If
        Return sql
    End Function

    Public Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        For i As Integer = 0 To _variables.Count - 1
            p.Add(_variables(i).getSqlParameter)
        Next
        Return p.ToArray()
    End Function

    Private Sub check()
        checkTable()
        checkVariables()
        MyBase.checkConnection()
    End Sub


End Class
