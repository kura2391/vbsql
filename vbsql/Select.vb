
Public Class [Select]
    Inherits SqlAbstract
    Private _select As New Parameter("select")
    Private _where As New Where("where")
    Private _order As String = Nothing


    Public Sub New(connection As Connection)
        MyBase.New(connection)
    End Sub

    Public Sub New(connectionString As String)
        MyBase.New(connectionString)
    End Sub

    'set table
    Public Function from(table As String, Optional ByVal type As SqlDbType = SqlDbType.NVarChar) As [Select]
        MyBase._table.Value = table
        MyBase._table.DbType = type
        Return Me
    End Function
    'set Columns
    Public Function [select](columns As String()) As [Select]
        _select.clear()
        For i As Integer = 0 To columns.Count - 1
            _select.appendParameter(columns(i), SqlDbType.NVarChar)
        Next
        Return Me
    End Function

    'set Where

    Public Function where(conditions As String, col() As String) As [Select]
        _where.set(conditions, col)
        Return Me
    End Function


    'set OrderBy
    Public Function orderBy(order As String) As [Select]
        _order = order
        Return Me
    End Function


    'execute sql and get data as DataTable
    Public Function execute(connectionString As String) As DataTable
        check()
        Return _connection.executeSelect(buildSql, buildParameter())
    End Function

    'create sql sentence 
    Private Function buildSql() As String
        Dim sql As String = "SELECT "

        For i As Integer = 0 To _select.count - 1
            sql &= "@" & _table.ParameterName & " ,"
        Next
        sql = sql.Substring(0, sql.Length - 1)

        sql &= " FROM @" & _table.ParameterName
        sql &= " "
        If Not _where.isEmpty() Then
            sql &= " WHERE "
            sql &= _where.sql()
        End If
        If Not IsNothing(_order) Then
            sql &= " ORDER BY "
            sql &= _order
        End If
        Return sql
    End Function

    Private Function buildParameter()
        Dim p As New List(Of SqlClient.SqlParameter)
        p.Add(_table)
        p.AddRange(_select.getParameterList)
        p.AddRange(_where.getParameterList())

        Return p
    End Function

    'if _from is not set, throw error
    Private Sub check()
        checkFrom()
        checkConnection()
    End Sub

    Private Sub checkFrom()
        If IsNothing(_table) OrElse _table.Value.Trim() = "" Then
            Throw New Exception("no table selected")
        End If
    End Sub
End Class