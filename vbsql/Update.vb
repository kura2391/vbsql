Public Class Update
    Inherits SqlAbstract

    Private _where As New Where()

    Public Sub New(connection As Connection)
        MyBase.New(connection)
        _prefix = "@set"
    End Sub
    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
        _prefix = "@set"
    End Sub


    Public Function table(tableName As String) As Update
        _table = tableName
        Return Me
    End Function

    Public Function [set](parameterList As List(Of Parameter)) As Update
        _variables.AddRange(parameterList)
        Return Me
    End Function

    'more easier set function
    Public Function [set](parameterList As Dictionary(Of String, String)) As Update
        Dim param As New List(Of Parameter)
        For Each key As String In parameterList.Keys
            param.Add(New Parameter(key, parameterList(key)))
        Next
        Return [set](param)
    End Function

    'set where
    Public Function where(conditions As String, col As Parameter()) As Update
        _where.add(conditions, col)
        Return Me
    End Function
    Public Function where(conditions As String, col() As String) As Update
        _where.add(conditions, col)
        Return Me
    End Function


    'execute sql and return integer 
    Public Function execute() As Integer
        check()
        setParameterName()
        Return _connection.execute(buildSql(), buildParameter())
    End Function

    '
    Function buildSql() As String
        Dim sql As String = "UPDATE " & _table & " SET "
        For i As Integer = 0 To _variables.Count() - 1 Step 1
            sql &= "" & _variables(i).getColumnName & " = " & _variables(i).getParameterName & ","
        Next
        sql = sql.Substring(0, sql.Length - 1)
        If Not _where.isEmpty() Then
            sql &= " WHERE " & _where.sql()
        End If
        Return sql
    End Function

    Private Function buildParameter() As SqlClient.SqlParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        For i As Integer = 0 To _variables.Count - 1
            p.Add(_variables(i).getSqlParameter)
        Next
        p.AddRange(_where.getParamList)

        Return p.ToArray()
    End Function


    Public Sub check()
        checkUpdate()
        checkConnection()
        checkTable()
    End Sub

    'if _from is not set, throw error
    Private Sub checkUpdate()
        If _table.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub


End Class
